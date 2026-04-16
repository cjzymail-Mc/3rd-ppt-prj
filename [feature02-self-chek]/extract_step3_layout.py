#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract per-slide element layout from Step3/deck_product_rnd.html.

Uses Playwright to render the HTML at 1260x720 CSS px and extract the
bounding box + computed style for every editable element in each slide.

Output: Step3/_layout_manifest.json
  {
    "canvas_w": 1260, "canvas_h": 720,
    "slides": [
      {
        "index": 1, "page": "01 / 16",
        "elements": [
          {"type":"text", "role":"hero-eyebrow", "text":"...",
           "x":56, "y":164, "w":900, "h":24,
           "font_size_px":13, "color":"#8c8c8c", "bold":false, "align":1},
          {"type":"image", "x":640, "y":430, "w":590, "h":248, "src":"images/img_01.jpg"},
          {"type":"table", "x":30, "y":95, "w":1200, "h":200,
           "rows":[[...]], "col_widths":[0.2, 0.4, 0.4]},
          ...
        ]
      }, ...
    ]
  }

element types:
  text        - plain text box (role: h1/h2/h3/p/note/hero-eyebrow/hero-subtitle/
                                hero-footer/priority-label/kpi-label-text/page-label/...)
  bullets     - <ul> bullet list (has "items" list)
  numbered    - <ol> numbered list (has "items" list)
  image       - <img> (has "src")
  table       - <table> (has "rows" 2-D list, "col_widths" fractions)
  kpi_card    - .kpi card (has "label" + "value" sub-texts, uses 12px inner padding)
  timeline    - .timeline-item row (has "badge" + "text")
  page_label  - slide page number (always last per slide)

Run:
  py -3 pipeline/extract_step3_layout.py
"""

import json
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT      = Path(r"D:\Technique Support\Claude Code Learning\[Agent-3 Codex] Info Classifier")
HTML      = ROOT / "Step3" / "deck_product_rnd.html"
OUT       = ROOT / "Step3" / "_layout_manifest.json"
CANVAS_W  = 1260
CANVAS_H  = 720

# Must match the layout normalization in screenshot_step3_masked.py so that
# extracted bounding boxes align exactly with the masked screenshots.
LAYOUT_CSS = """
.wrapper {
    padding: 0 !important;
    max-width: 1260px !important;
    margin: 0 !important;
}
.slide {
    width: 1260px !important;
    min-height: 720px !important;
    max-height: 720px !important;
    height: 720px !important;
    margin: 0 !important;
    border-radius: 0 !important;
    overflow: hidden !important;
    box-shadow: none !important;
}
"""

# ---------------------------------------------------------------------------
# JavaScript extractor (runs inside the browser page)
# ---------------------------------------------------------------------------

_EXTRACTOR_JS = """(slide) => {
    const slideBB = slide.getBoundingClientRect();
    const elements = [];
    const processed = new WeakSet();

    function relBB(el) {
        const bb = el.getBoundingClientRect();
        return {
            x: Math.round(bb.left - slideBB.left),
            y: Math.round(bb.top  - slideBB.top),
            w: Math.round(bb.width),
            h: Math.round(bb.height)
        };
    }

    function rgbToHex(c) {
        if (!c || c === 'transparent') return null;
        const m = c.match(/[\\d.]+/g);
        if (!m || m.length < 3) return null;
        let r = parseInt(m[0]), g = parseInt(m[1]), b = parseInt(m[2]);
        const a = m.length >= 4 ? parseFloat(m[3]) : 1.0;
        // Blend semi-transparent colors against the effective background.
        // Hero slide (dark bg ~#0d1424) vs regular slides (white bg #ffffff).
        if (a < 0.99) {
            const el = arguments[1];  // optional context element
            const isHero = el && el.closest && el.closest('.slide-hero');
            const bgR = isHero ? 13 : 255;
            const bgG = isHero ? 20 : 255;
            const bgB = isHero ? 36 : 255;
            r = Math.round(r * a + bgR * (1 - a));
            g = Math.round(g * a + bgG * (1 - a));
            b = Math.round(b * a + bgB * (1 - a));
        }
        return '#' + [r,g,b].map(n => n.toString(16).padStart(2,'0')).join('');
    }

    function getStyle(el) {
        const cs = window.getComputedStyle(el);
        const fw = cs.fontWeight;
        const bold = parseInt(fw) >= 600 || fw === 'bold' || fw === 'bolder';
        const ta = cs.textAlign;
        const align = ta === 'center' ? 2 : (ta === 'right' || ta === 'end') ? 3 : 1;
        const color = rgbToHex(cs.color, el) || '#111827';
        return {
            font_size_px: Math.round(parseFloat(cs.fontSize)),
            color,
            bold,
            align,
            pad_l: Math.round(parseFloat(cs.paddingLeft) || 0),
            pad_t: Math.round(parseFloat(cs.paddingTop) || 0),
            pad_r: Math.round(parseFloat(cs.paddingRight) || 0),
            pad_b: Math.round(parseFloat(cs.paddingBottom) || 0),
        };
    }

    function getPadding(el) {
        const cs = window.getComputedStyle(el);
        return {
            pad_l: Math.round(parseFloat(cs.paddingLeft) || 0),
            pad_t: Math.round(parseFloat(cs.paddingTop) || 0),
            pad_r: Math.round(parseFloat(cs.paddingRight) || 0),
            pad_b: Math.round(parseFloat(cs.paddingBottom) || 0),
        };
    }

    function getText(el) {
        return (el.innerText || '').trim();
    }

    function mark(el) { processed.add(el); }
    function seen(el) { return processed.has(el); }

    // ── 1. Tables ────────────────────────────────────────────────────────
    for (const tbl of slide.querySelectorAll('table')) {
        if (seen(tbl)) continue;
        mark(tbl);
        const pos = relBB(tbl);
        if (pos.w < 2) continue;
        const rows = [];
        const colWidths = [];
        let firstRow = true;
        for (const tr of tbl.querySelectorAll('tr')) {
            mark(tr);
            const cells = tr.querySelectorAll('th,td');
            const row = [];
            for (const cell of cells) {
                mark(cell);
                const cs = window.getComputedStyle(cell);
                const fw = cs.fontWeight;
                if (firstRow) colWidths.push(Math.round(cell.getBoundingClientRect().width));
                row.push({
                    text: getText(cell),
                    is_header: cell.tagName === 'TH',
                    align: cs.textAlign === 'center' ? 2 : cs.textAlign === 'right' ? 3 : 1,
                    bold: parseInt(fw) >= 600 || fw === 'bold',
                    colspan: cell.colSpan || 1
                });
            }
            if (row.length) rows.push(row);
            firstRow = false;
        }
        const totalW = colWidths.reduce((a,b)=>a+b, 0) || 1;
        elements.push({
            type: 'table', ...pos, rows,
            col_widths: colWidths.map(w => parseFloat((w / totalW).toFixed(4)))
        });
    }

    // ── 2. Images ────────────────────────────────────────────────────────
    for (const img of slide.querySelectorAll('img')) {
        if (seen(img)) continue;
        mark(img);
        const pos = relBB(img);
        if (pos.w < 5) continue;
        elements.push({type:'image', ...pos, src: img.getAttribute('src') || ''});
    }

    // ── 3. Lists (ul/ol) not inside tables ───────────────────────────────
    for (const lst of slide.querySelectorAll('ul,ol')) {
        if (seen(lst)) continue;
        if (lst.closest('table')) continue;
        mark(lst);
        const items = [];
        for (const li of lst.querySelectorAll('li')) {
            mark(li);
            items.push(getText(li));
        }
        if (!items.length) continue;
        const pos = relBB(lst);
        if (pos.w < 2) continue;
        const firstLi = lst.querySelector('li');
        const cs = firstLi ? window.getComputedStyle(firstLi) : null;
        const color = cs ? (rgbToHex(cs.color) || '#111827') : '#111827';
        const listPad = getPadding(lst);
        elements.push({
            type: lst.tagName === 'OL' ? 'numbered' : 'bullets',
            ...pos, items,
            font_size_px: cs ? Math.round(parseFloat(cs.fontSize)) : 15,
            color,
            ...listPad
        });
    }

    // ── 4. Headings ──────────────────────────────────────────────────────
    for (const tag of ['h1','h2','h3']) {
        for (const h of slide.querySelectorAll(tag)) {
            if (seen(h)) continue;
            mark(h);
            const pos = relBB(h);
            if (pos.w < 2) continue;
            elements.push({type:'text', role:tag, text:getText(h), ...pos, ...getStyle(h)});
        }
    }

    // ── 5. Named role elements ────────────────────────────────────────────
    const roleMap = {
        '.note':           'note',
        '.hero-eyebrow':   'hero-eyebrow',
        '.hero-subtitle':  'hero-subtitle',
        '.hero-footer':    'hero-footer',
    };
    for (const [sel, role] of Object.entries(roleMap)) {
        for (const el of slide.querySelectorAll(sel)) {
            if (seen(el)) continue;
            mark(el);
            const pos = relBB(el);
            if (pos.w < 2) continue;
            elements.push({type:'text', role, text:getText(el), ...pos, ...getStyle(el)});
        }
    }

    // ── 6. Hero KPI items (val + label pairs) ────────────────────────────
    for (const item of slide.querySelectorAll('.hero-kpi-item')) {
        if (seen(item)) continue;
        mark(item);
        const valEl  = item.querySelector('.hero-kpi-val');
        const lblEl  = item.querySelector('.hero-kpi-label');
        for (const [el, role] of [[valEl,'kpi-val'],[lblEl,'kpi-label']]) {
            if (!el || seen(el)) continue;
            mark(el);
            const pos = relBB(el);
            if (pos.w < 2) continue;
            elements.push({type:'text', role, text:getText(el), ...pos, ...getStyle(el)});
        }
    }

    // ── 7. KPI cards (.kpi) ──────────────────────────────────────────────
    for (const kpi of slide.querySelectorAll('.kpi')) {
        if (seen(kpi)) continue;
        mark(kpi);
        const pos = relBB(kpi);
        if (pos.w < 2) continue;
        // Label = direct text nodes; Value = <b> child
        let labelText = '';
        for (const node of kpi.childNodes) {
            if (node.nodeType === 3) labelText += node.textContent;
        }
        labelText = labelText.trim();
        const bEl = kpi.querySelector('b');
        const valText = bEl ? getText(bEl) : '';
        if (bEl) mark(bEl);
        elements.push({
            type: 'kpi_card', ...pos,
            label: labelText, value: valText
        });
    }

    // ── 8. Timeline items (.timeline-item) ───────────────────────────────
    for (const item of slide.querySelectorAll('.timeline-item')) {
        if (seen(item)) continue;
        mark(item);
        const pos = relBB(item);
        if (pos.w < 2) continue;
        const badge = item.querySelector('.timeline-badge');
        const badgeText = badge ? getText(badge) : '';
        if (badge) mark(badge);
        // Content = all non-badge span text
        let contentText = '';
        for (const span of item.querySelectorAll('span:not(.timeline-badge)')) {
            mark(span);
            contentText += getText(span);
        }
        elements.push({type:'timeline', ...pos, badge:badgeText, text:contentText.trim()});
    }

    // ── 9. Priority blocks (.priority-p1/p2/p3) — mark so p tags inside
    //       are not double-extracted; the ol inside was already captured above
    for (const block of slide.querySelectorAll('.priority-p1,.priority-p2,.priority-p3')) {
        if (seen(block)) continue;
        mark(block);
        const lbl = block.querySelector('.priority-label');
        if (lbl && !seen(lbl)) {
            mark(lbl);
            const pos = relBB(lbl);
            if (pos.w >= 2) {
                elements.push({
                    type:'text', role:'priority-label',
                    text:getText(lbl), ...pos, ...getStyle(lbl)
                });
            }
        }
    }

    // ── 10. Tags (.tag) in header area ───────────────────────────────────
    for (const tag of slide.querySelectorAll('.tag')) {
        if (seen(tag)) continue;
        mark(tag);
        // .tag elements are inline decorations in a <p> — extract the parent p
    }

    // ── 11. Remaining standalone <p> elements ────────────────────────────
    for (const p of slide.querySelectorAll('p')) {
        if (seen(p)) continue;
        if (p.closest('ul,ol,table,.note,.hero-eyebrow,.hero-subtitle,.hero-footer,.timeline-item,.priority-p1,.priority-p2,.priority-p3')) continue;
        mark(p);
        const text = getText(p);
        if (!text) continue;
        const pos = relBB(p);
        if (pos.w < 2) continue;
        elements.push({type:'text', role:'p', text, ...pos, ...getStyle(p)});
    }

    // ── 12. Page label (from data-page attribute) ─────────────────────────
    const pageAttr = slide.getAttribute('data-page') || '';
    // Page label text for slide-hero is derived from its number (01/16)
    // For all slides, position = bottom-right, right:20px, bottom:14px, ~152x18
    const pageText = pageAttr;
    if (pageText) {
        elements.push({
            type: 'page_label',
            x: 1260 - 20 - 152, y: 720 - 14 - 18,
            w: 152, h: 18,
            text: pageText
        });
    }

    return elements;
}"""


def main():
    print("Starting Playwright ...")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(
            viewport={"width": CANVAS_W, "height": CANVAS_H}
        )
        print(f"Loading {HTML.name} ...")
        page.goto(HTML.as_uri(), wait_until="domcontentloaded")
        page.wait_for_timeout(800)
        # Apply the same layout normalization used in screenshot_step3_masked.py
        # so extracted bounding boxes match the screenshot coordinate space.
        page.add_style_tag(content=LAYOUT_CSS)
        page.wait_for_timeout(300)  # let layout reflow settle

        slides = page.query_selector_all("section.slide")
        n = len(slides)
        print(f"Found {n} slides (expected 16)")

        manifest = {
            "canvas_w": CANVAS_W,
            "canvas_h": CANVAS_H,
            "slides": [],
        }

        for i, slide in enumerate(slides, start=1):
            page_attr = page.evaluate("el => el.getAttribute('data-page') || ''", slide)
            print(f"  Slide {i:02d} ({page_attr or 'hero'}) ...", end="", flush=True)
            elements = page.evaluate(_EXTRACTOR_JS, slide)
            # Hero slide page label not in data-page — inject manually
            if i == 1 and not any(e["type"] == "page_label" for e in elements):
                elements.append({
                    "type": "page_label",
                    "x": CANVAS_W - 20 - 152,
                    "y": CANVAS_H - 14 - 18,
                    "w": 152, "h": 18,
                    "text": "01 / 16",
                })
            manifest["slides"].append({
                "index": i,
                "page": page_attr or f"{i:02d} / 16",
                "elements": elements,
            })
            counts = {}
            for e in elements:
                counts[e["type"]] = counts.get(e["type"], 0) + 1
            summary = ", ".join(f"{v} {k}" for k, v in sorted(counts.items()))
            print(f" {len(elements)} elements ({summary})")

        browser.close()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"\nManifest saved → {OUT}")
    print(f"Total: {n} slides extracted")


if __name__ == "__main__":
    main()
