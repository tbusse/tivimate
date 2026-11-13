#!/usr/bin/env python3
"""
PPTX Presentation Translator (EN → DE)

Translates an English PowerPoint presentation to German while preserving
all design, layout, and formatting. Automatically adjusts font sizes
when German text is significantly longer than the English original.

Usage:
    # With Anthropic API (recommended — best quality, layout-aware):
    export ANTHROPIC_API_KEY=sk-ant-...
    python translate_pptx.py presentation.pptx

    # With explicit output path:
    python translate_pptx.py input.pptx -o output_de.pptx

    # With Google Translate fallback (no API key needed):
    python translate_pptx.py input.pptx --google

Requires:
    pip install python-pptx anthropic
    pip install deep-translator   # only needed with --google
"""

import argparse
import copy
import json
import os
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt, Emu
from lxml import etree


# ---------------------------------------------------------------------------
# Text collection
# ---------------------------------------------------------------------------

def collect_text_segments(prs):
    """Walk every slide and collect translatable text segments with metadata."""
    segments = []

    for slide_idx, slide in enumerate(prs.slides):
        _collect_from_shapes(slide.shapes, segments, slide_idx)

        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            tf = slide.notes_slide.notes_text_frame
            full = tf.text.strip()
            if full and not _is_placeholder_notes(full):
                segments.append({
                    "type": "notes",
                    "slide": slide_idx,
                    "text": full,
                    "text_frame": tf,
                })

    return segments


def _is_placeholder_notes(text):
    lower = text.strip().lower()
    return lower in (
        "click to add notes",
        "click to add text",
        "klicken sie, um notizen hinzuzufügen",
    )


def _collect_from_shapes(shapes, segments, slide_idx):
    for shape in shapes:
        if shape.shape_type == 6:  # group shape
            _collect_from_shapes(shape.shapes, segments, slide_idx)
            continue

        if shape.has_table:
            table = shape.table
            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    text = cell.text_frame.text.strip()
                    if text:
                        segments.append({
                            "type": "table_cell",
                            "slide": slide_idx,
                            "row": row_idx,
                            "col": col_idx,
                            "text": text,
                            "text_frame": cell.text_frame,
                        })
            continue

        if shape.has_text_frame:
            tf = shape.text_frame
            full_text = tf.text.strip()
            if full_text:
                segments.append({
                    "type": "text",
                    "slide": slide_idx,
                    "shape_name": shape.name,
                    "text": full_text,
                    "text_frame": tf,
                    "shape": shape,
                })


# ---------------------------------------------------------------------------
# Translation engines
# ---------------------------------------------------------------------------

def translate_batch_anthropic(texts, api_key, model="claude-sonnet-4-20250514"):
    """Translate a batch of texts using the Anthropic API."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)

    numbered = "\n".join(f"[{i}] {t}" for i, t in enumerate(texts))

    response = client.messages.create(
        model=model,
        max_tokens=8192,
        messages=[{
            "role": "user",
            "content": (
                "Übersetze die folgenden englischen Texte ins Deutsche.\n"
                "Diese Texte stammen aus einer PowerPoint-Präsentation.\n\n"
                "Wichtige Regeln:\n"
                "- Halte die Übersetzungen möglichst kurz und prägnant "
                "(Slides haben wenig Platz)\n"
                "- Wenn der deutsche Text deutlich länger wäre als der "
                "englische, kürze ihn intelligent\n"
                "- Behalte Fachbegriffe bei, die im Deutschen üblich auf "
                "Englisch verwendet werden\n"
                "- Behalte Eigennamen, Produktnamen, URLs, E-Mail-Adressen "
                "und Zahlen unverändert\n"
                "- Bullet Points und Aufzählungen als solche beibehalten\n"
                "- Zeilenumbrüche (\\n) innerhalb eines Textes beibehalten\n"
                "- Formale Anrede verwenden (Sie statt Du)\n"
                "- Antworte NUR mit einem JSON-Array von Strings, in der "
                "gleichen Reihenfolge\n\n"
                f"Texte:\n{numbered}\n\n"
                'Antworte ausschließlich mit einem JSON-Array, z.B.: '
                '["Text 1", "Text 2", ...]'
            ),
        }],
    )

    content = response.content[0].text.strip()
    match = re.search(r'\[.*\]', content, re.DOTALL)
    if match:
        return json.loads(match.group())
    return json.loads(content)


def translate_batch_google(texts):
    """Fallback: translate using Google Translate (no API key needed)."""
    from deep_translator import GoogleTranslator

    translator = GoogleTranslator(source='en', target='de')
    results = []
    for text in texts:
        try:
            translated = translator.translate(text)
            results.append(translated if translated else text)
        except Exception as e:
            print(f"  ⚠️  Google Translate Fehler: {e}")
            results.append(text)
    return results


# ---------------------------------------------------------------------------
# Text replacement with formatting preservation
# ---------------------------------------------------------------------------

def apply_translation(text_frame, new_text):
    """Replace text in a text_frame while preserving per-run formatting.

    Strategy: snapshot all paragraph/run formatting, clear the XML body,
    rebuild paragraphs with new text and old formatting.
    """
    # 1. Snapshot formatting for every paragraph + run
    para_snapshots = []
    for para in text_frame.paragraphs:
        pxml = copy.deepcopy(para._p)
        runs_data = []
        for run in para.runs:
            runs_data.append({
                "text": run.text,
                "font": _copy_font_props(run.font),
            })
        para_snapshots.append({
            "alignment": para.alignment,
            "level": para.level,
            "space_before": para.space_before,
            "space_after": para.space_after,
            "line_spacing": para.line_spacing,
            "runs": runs_data,
            "pxml": pxml,
        })

    new_lines = new_text.split('\n')

    # 2. Match paragraph count to new line count
    while len(para_snapshots) > len(new_lines):
        para_snapshots.pop()
    while len(new_lines) > len(para_snapshots):
        para_snapshots.append(copy.deepcopy(para_snapshots[-1]))

    # 3. Remove ALL existing <a:p> elements from the text body XML
    txBody = text_frame._txBody
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    for existing_p in txBody.findall('a:p', nsmap):
        txBody.remove(existing_p)

    # 4. Rebuild each paragraph
    for line, pdata in zip(new_lines, para_snapshots):
        # Clone the original paragraph XML as a structural template
        new_p = copy.deepcopy(pdata["pxml"])

        # Remove all <a:r> (run) and <a:br> (break) elements from template
        for child in list(new_p):
            tag = etree.QName(child.tag).localname if isinstance(child.tag, str) else ''
            if tag in ('r', 'br'):
                new_p.remove(child)

        # Create new runs with translated text
        if not pdata["runs"]:
            _add_run_xml(new_p, line, {})
        elif len(pdata["runs"]) == 1:
            _add_run_xml(new_p, line, pdata["runs"][0]["font"])
        else:
            _distribute_text_to_runs_xml(new_p, line, pdata["runs"])

        txBody.append(new_p)


def _add_run_xml(p_elem, text, font_props):
    """Add an <a:r> element to a paragraph element with formatting."""
    nsmap = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    r = etree.SubElement(p_elem, f'{{{nsmap}}}r')
    rPr = etree.SubElement(r, f'{{{nsmap}}}rPr', attrib={'lang': 'de-DE'})

    if font_props.get("bold"):
        rPr.set('b', '1')
    if font_props.get("italic"):
        rPr.set('i', '1')
    if font_props.get("underline"):
        rPr.set('u', 'sng')
    if font_props.get("size") is not None:
        rPr.set('sz', str(font_props["size"] // 12700))  # EMU to hundredths of pt

    if "color_rgb" in font_props:
        solidFill = etree.SubElement(rPr, f'{{{nsmap}}}solidFill')
        srgbClr = etree.SubElement(solidFill, f'{{{nsmap}}}srgbClr')
        srgbClr.set('val', str(font_props["color_rgb"]))

    if "name" in font_props:
        latin = etree.SubElement(rPr, f'{{{nsmap}}}latin')
        latin.set('typeface', font_props["name"])

    t = etree.SubElement(r, f'{{{nsmap}}}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return r


def _distribute_text_to_runs_xml(p_elem, text, runs_data):
    """Distribute translated text proportionally across multiple styled runs."""
    old_total = sum(len(r["text"]) for r in runs_data)
    if old_total == 0:
        _add_run_xml(p_elem, text, runs_data[0]["font"])
        return

    remaining = text
    for j, rdata in enumerate(runs_data):
        if j == len(runs_data) - 1:
            _add_run_xml(p_elem, remaining, rdata["font"])
        else:
            proportion = len(rdata["text"]) / old_total
            split_pos = max(1, int(len(text) * proportion))
            # try to split at a word boundary
            space_pos = remaining.find(' ', max(0, split_pos - 5))
            if 0 < space_pos <= split_pos + 8:
                split_pos = space_pos + 1
            _add_run_xml(p_elem, remaining[:split_pos], rdata["font"])
            remaining = remaining[split_pos:]


def _copy_font_props(font):
    """Snapshot font properties."""
    props = {}
    for attr in ("bold", "italic", "underline", "size"):
        try:
            val = getattr(font, attr)
            if val is not None:
                props[attr] = val
        except Exception:
            pass
    try:
        if font.color and font.color.type is not None:
            props["color_rgb"] = font.color.rgb
    except Exception:
        pass
    try:
        if font.name:
            props["name"] = font.name
    except Exception:
        pass
    return props


# ---------------------------------------------------------------------------
# Font-size adjustment
# ---------------------------------------------------------------------------

def adjust_font_sizes(segments, translations):
    """Reduce font size proportionally when translation is significantly longer."""
    THRESHOLD = 1.2

    for seg, new_text in zip(segments, translations):
        old_len = len(seg["text"])
        new_len = len(new_text)

        if old_len == 0 or new_len <= old_len * THRESHOLD:
            continue

        ratio = old_len / new_len
        scale = max(ratio, 0.70)

        tf = seg["text_frame"]
        nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
        for rPr in tf._txBody.findall('.//a:rPr', nsmap):
            sz = rPr.get('sz')
            if sz:
                new_sz = max(800, int(int(sz) * scale))
                rPr.set('sz', str(new_sz))


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------

def translate_presentation(input_path, output_path, api_key=None,
                           model="claude-sonnet-4-20250514",
                           use_google=False, verbose=True):
    """Main translation pipeline."""
    if verbose:
        print(f"Lade Praesentation: {input_path}")

    prs = Presentation(input_path)
    segments = collect_text_segments(prs)

    if not segments:
        print("Keine uebersetzbaren Texte gefunden.")
        return

    texts = [s["text"] for s in segments]

    if verbose:
        print(f"{len(texts)} Textsegmente auf {len(prs.slides)} Folien gefunden")

    BATCH_SIZE = 40
    all_translations = []

    for i in range(0, len(texts), BATCH_SIZE):
        batch = texts[i:i + BATCH_SIZE]
        batch_num = i // BATCH_SIZE + 1
        total_batches = (len(texts) + BATCH_SIZE - 1) // BATCH_SIZE

        if verbose:
            print(f"Uebersetze Batch {batch_num}/{total_batches} "
                  f"({len(batch)} Segmente)...")

        if use_google:
            if verbose:
                print("   -> Google Translate")
            translated = translate_batch_google(batch)
        elif api_key:
            try:
                if verbose:
                    print(f"   -> Anthropic API ({model})")
                translated = translate_batch_anthropic(batch, api_key, model)
            except Exception as e:
                print(f"  Anthropic API fehlgeschlagen: {e}")
                print("  Fallback auf Google Translate...")
                translated = translate_batch_google(batch)
        else:
            print("FEHLER: Kein ANTHROPIC_API_KEY gesetzt und --google nicht angegeben.")
            print()
            print("Optionen:")
            print("  export ANTHROPIC_API_KEY=sk-ant-...")
            print("  python translate_pptx.py input.pptx")
            print()
            print("  python translate_pptx.py input.pptx --google")
            sys.exit(1)

        if len(translated) != len(batch):
            print(f"  Erwartete {len(batch)} Uebersetzungen, "
                  f"bekam {len(translated)}. Ergaenze fehlende...")
            while len(translated) < len(batch):
                translated.append(batch[len(translated)])

        all_translations.extend(translated)

    if verbose:
        print("Ersetze Texte und passe Layout an...")

    for seg, new_text in zip(segments, all_translations):
        apply_translation(seg["text_frame"], new_text)

    adjust_font_sizes(segments, all_translations)

    prs.save(output_path)

    if verbose:
        print(f"Deutsche Version gespeichert: {output_path}")
        print()
        _print_summary(segments, all_translations)


def _print_summary(segments, translations):
    slides = set(s["slide"] for s in segments)
    print(f"   Folien bearbeitet: {len(slides)}")
    print(f"   Textsegmente uebersetzt: {len(translations)}")

    longer = sum(1 for s, t in zip(segments, translations)
                 if len(t) > len(s["text"]) * 1.2)
    if longer:
        print(f"   Schriftgroesse angepasst: {longer} Segmente")


def main():
    parser = argparse.ArgumentParser(
        description="Uebersetze eine englische PowerPoint-Praesentation ins Deutsche"
    )
    parser.add_argument("input", help="Pfad zur englischen PPTX-Datei")
    parser.add_argument("output", nargs="?", help="Pfad fuer die deutsche PPTX-Datei")
    parser.add_argument("-o", "--output-file", help="Pfad fuer die deutsche PPTX-Datei")
    parser.add_argument(
        "--api-key",
        default=os.environ.get("ANTHROPIC_API_KEY"),
        help="Anthropic API Key (oder ANTHROPIC_API_KEY Umgebungsvariable)",
    )
    parser.add_argument(
        "--model",
        default="claude-sonnet-4-20250514",
        help="Anthropic-Modell fuer die Uebersetzung",
    )
    parser.add_argument(
        "--google", action="store_true",
        help="Google Translate statt Anthropic API verwenden (kein API-Key noetig)",
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true",
        help="Keine Fortschrittsausgabe",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Fehler: Datei nicht gefunden: {input_path}")
        sys.exit(1)

    output_path = args.output_file or args.output
    if not output_path:
        output_path = input_path.stem + "_DE" + input_path.suffix
    output_path = Path(output_path)

    translate_presentation(
        str(input_path),
        str(output_path),
        api_key=args.api_key,
        model=args.model,
        use_google=args.google,
        verbose=not args.quiet,
    )


if __name__ == "__main__":
    main()
