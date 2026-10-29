"""Convert your deck: someone uploads their own PowerPoint or PDF and gets it back rebuilt on the evoila template.

Fired for Deck Requests rows with Request type = "Convert deck" (the hub's Convert your deck page). The file is
on the row's "Deck to convert" field; the page's options (deck title, wording mode, a download link instead of
a file) are JSON in the Custom outline field.

1. read_source(): renders every slide to an image and pulls out what the slide is made of: titles and text,
   native charts with their numbers, tables, SmartArt text, pictures and drawn diagrams with their position.
2. map_slides(): Claude looks at each slide (the picture and the extracted parts) and picks the evoila layout
   that carries the same idea: bullets become bullets or icon cards, chevrons and arrows become steps, a
   radial diagram becomes the hub layout, native charts are redrawn in evoila colours from their real numbers,
   tables become evoila tables, and visuals that cannot be redrawn faithfully (screenshots, architecture
   drawings, partner logos, maps) are kept as a framed image cropped from the original slide. Decorative
   stock photos and the old template's logos and backgrounds are dropped.
3. custom_deck.render() lays the result out on the master template; the PPTX and PDF go back to the row and
   the existing "Deck ready" automation emails them. The Market data check field gets a slide-by-slide report
   of what each original slide became.

Nothing here prints request data: the repo is public and the build logs are visible.
"""
import base64, concurrent.futures as cf, datetime, io, json, os, re, subprocess, sys, urllib.request, zipfile

import run_request as rr
import ai_steps
import custom_deck as CD

F_SOURCE = "fldjNIm3iace0MiLH"
MAX_SLIDES = 40
BATCH = 5
EMU_W, EMU_H = 12192000, 6858000


# ---------------------------------------------------------------- 1. read the source deck

def download_source(f, opts):
    """Return the local path of the uploaded deck (pptx or pdf)."""
    att = f.get(F_SOURCE) or []
    url, name = (att[0]["url"], att[0].get("filename") or "deck") if att else (opts.get("link") or "", "deck")
    if not url:
        raise RuntimeError("No deck was attached to this request.")
    url = drive_direct(url)
    raw = os.path.join(rr.WORK, "source.bin")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 evoila-deck-converter"})
    with urllib.request.urlopen(req, timeout=300) as r, open(raw, "wb") as fh:
        fh.write(r.read())
    with open(raw, "rb") as fh:
        head = fh.read(8)
    if head.startswith(b"%PDF"):
        kind = "pdf"
    elif head.startswith(b"PK"):
        with zipfile.ZipFile(raw) as z:
            names = z.namelist()
        if not any(n.startswith("ppt/") for n in names):
            raise RuntimeError("The file is not a PowerPoint deck. Upload a .pptx or a PDF.")
        kind = "pptx"
    elif head.startswith(b"\xd0\xcf\x11\xe0"):
        kind = "ppt"                           # old binary PowerPoint: LibreOffice converts it first
    else:
        raise RuntimeError("Could not open the deck. If you used a link, it must be a direct download link "
                           "that works without signing in. Otherwise upload the .pptx or PDF.")
    path = os.path.join(rr.WORK, "source." + kind)
    os.replace(raw, path)
    if kind == "ppt":
        subprocess.run(["soffice", "--headless", "--convert-to", "pptx", "--outdir", rr.WORK, path],
                       check=True, timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        path, kind = os.path.join(rr.WORK, "source.pptx"), "pptx"
    return path, kind, name


def drive_direct(url):
    """Google Drive and Dropbox share links -> direct downloads. Other links are used as given."""
    m = re.search(r"drive\.google\.com/file/d/([\w-]+)", url) or re.search(r"drive\.google\.com/open\?id=([\w-]+)", url)
    if m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    m = re.search(r"docs\.google\.com/presentation/d/([\w-]+)", url)
    if m:
        return f"https://docs.google.com/presentation/d/{m.group(1)}/export/pptx"
    if "dropbox.com" in url:
        return re.sub(r"([?&])dl=0", r"\1dl=1", url) if "dl=0" in url else url + ("&" if "?" in url else "?") + "dl=1"
    return url


def render_pages(path, kind):
    """PDF of the deck plus one PNG per page at 150 dpi (for Claude and for cropping kept visuals)."""
    pdf = path if kind == "pdf" else None
    if not pdf:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", rr.WORK, path],
                       check=True, timeout=300, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = os.path.splitext(path)[0] + ".pdf"
    out = os.path.join(rr.WORK, "page")
    subprocess.run(["pdftoppm", "-r", "150", "-png", "-l", str(MAX_SLIDES + 5), pdf, out], check=True, timeout=300)
    pages = sorted(x for x in os.listdir(rr.WORK) if x.startswith("page-") and x.endswith(".png"))
    return [os.path.join(rr.WORK, p) for p in pages], pdf


def _bbox(sh, ox=0, oy=0):
    try:
        return [round((sh.left + ox) / EMU_W, 3), round((sh.top + oy) / EMU_H, 3),
                round(sh.width / EMU_W, 3), round(sh.height / EMU_H, 3)]
    except Exception:
        return None


def _smartart_text(slide, gf):
    """SmartArt keeps its text in a separate diagram data part."""
    try:
        from lxml import etree
        rid = gf._element.xpath(".//*[local-name()='relIds']/@*[local-name()='dm']")
        if not rid:
            return []
        part = slide.part.related_part(rid[0])
        root = etree.fromstring(part.blob)
        return [t.text.strip() for t in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")
                if t.text and t.text.strip()]
    except Exception:
        return []


def extract_pptx(path):
    """Per visible slide: title, text blocks, tables, chart data, visuals (pictures, drawn diagrams), notes."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
    global EMU_W, EMU_H
    prs = Presentation(path)
    EMU_W, EMU_H = prs.slide_width or EMU_W, prs.slide_height or EMU_H
    slides, page = [], 0
    for idx, s in enumerate(prs.slides, 1):
        if s._element.get("show") == "0":
            continue                            # hidden slides are not in the rendered PDF either
        page += 1
        d = {"n": page, "title": "", "text": [], "tables": [], "charts": [], "visuals": [], "notes": ""}
        vid = [0]

        def visual(kind, sh, desc, ox=0, oy=0):
            bb = _bbox(sh, ox, oy)
            if not bb or bb[2] * bb[3] < 0.012:   # icons and bullets, not worth keeping
                return
            vid[0] += 1
            d["visuals"].append({"id": f"V{vid[0]}", "kind": kind, "box": bb, "about": desc[:160]})

        def walk(shapes, ox=0, oy=0, depth=0):
            for sh in shapes:
                try:
                    st = sh.shape_type
                except Exception:
                    st = None
                if st == MSO_SHAPE_TYPE.GROUP:
                    inner = []
                    for x in sh.shapes:
                        if getattr(x, "has_text_frame", False) and x.text_frame.text.strip():
                            inner.append(x.text_frame.text.strip().replace("\n", " / "))
                    if depth == 0 and len(list(sh.shapes)) >= 3:
                        visual("drawn diagram (group)", sh, "; ".join(inner))
                    walk(sh.shapes, ox, oy, depth + 1)
                    continue
                if getattr(sh, "has_chart", False) and sh.has_chart:
                    ch = sh.chart
                    try:
                        cats = [str(c) for c in ch.plots[0].categories]
                    except Exception:
                        cats = []
                    series = []
                    for pl in ch.plots:
                        for se in pl.series:
                            try:
                                series.append({"name": se.name or "", "values": [v if v is not None else 0 for v in se.values]})
                            except Exception:
                                pass
                    cid = f"C{len(d['charts']) + 1}"
                    title = ""
                    try:
                        if ch.has_title:
                            title = ch.chart_title.text_frame.text
                    except Exception:
                        pass
                    d["charts"].append({"id": cid, "type": str(ch.chart_type).split(".")[-1].split(" ")[0],
                                        "title": title, "categories": cats, "series": series})
                    visual("native chart " + cid, sh, title or "chart", ox, oy)
                    continue
                if getattr(sh, "has_table", False) and sh.has_table:
                    d["tables"].append([[c.text.strip() for c in r.cells] for r in sh.table.rows])
                    continue
                if st == MSO_SHAPE_TYPE.PICTURE or (st == MSO_SHAPE_TYPE.PLACEHOLDER and hasattr(sh, "image")):
                    alt = ""
                    try:
                        alt = sh._element.xpath("./p:nvPicPr/p:cNvPr/@descr")[0]
                    except Exception:
                        pass
                    visual("picture", sh, alt or "picture", ox, oy)
                    continue
                if sh._element.tag.endswith("}graphicFrame"):
                    txt = _smartart_text(s, sh)
                    if txt:
                        d["text"].append("[SmartArt] " + " | ".join(txt))
                        visual("SmartArt diagram", sh, " | ".join(txt), ox, oy)
                    continue
                if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                    t = sh.text_frame.text.strip()
                    is_title = False
                    if sh.is_placeholder:
                        try:
                            is_title = sh.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
                        except Exception:
                            pass
                    if is_title and not d["title"]:
                        d["title"] = t
                    else:
                        auto = ""
                        try:
                            if st == MSO_SHAPE_TYPE.AUTO_SHAPE:
                                auto = f"[{str(sh.auto_shape_type).split('.')[-1].split(' ')[0].lower()} shape] "
                        except Exception:
                            pass
                        d["text"].append(auto + t)
        walk(s.shapes)
        try:
            if s.has_notes_slide:
                d["notes"] = s.notes_slide.notes_text_frame.text.strip()[:1500]
        except Exception:
            pass
        slides.append(d)
    return slides


def extract_pdf(pdf, n_pages):
    out = []
    for i in range(1, n_pages + 1):
        try:
            t = subprocess.run(["pdftotext", "-f", str(i), "-l", str(i), "-layout", pdf, "-"],
                               capture_output=True, text=True, timeout=60).stdout
        except Exception:
            t = ""
        out.append({"n": i, "title": "", "text": [re.sub(r"[ \t]{3,}", "   ", t.strip())[:4000]],
                    "tables": [], "charts": [], "visuals": [], "notes": ""})
    return out


# ---------------------------------------------------------------- 2. Claude maps each slide

LAYOUTS = CD.LAYOUTS + ["chart", "table", "image", "hub"]
SLIDE_SCHEMA = {"type": "object", "properties": {
    "source_slide": {"type": "number"},
    "layout": {"type": "string", "enum": LAYOUTS},
    "title": {"type": "string"}, "kicker": {"type": "string"},
    "statement": {"type": "string"}, "support": {"type": "string"},
    "items": {"type": "array", "items": CD.ITEM},
    "headers": {"type": "array", "items": {"type": "string"}},
    "rows": {"type": "array", "items": {"type": "object", "properties": {
        "label": {"type": "string"}, "cells": {"type": "array", "items": {"type": "string"}}}}},
    "chart_ref": {"type": "string"}, "chart_kind": {"type": "string", "enum": ["column", "bar", "line", "pie"]},
    "visual": {"type": "string"},
    "crop": {"type": "array", "items": {"type": "number"}},
    "caption": {"type": "string"},
    "hub_center": {"type": "string"},
    "source": {"type": "string"}, "notes": {"type": "string"},
    "change": {"type": "string"}},
    "required": ["source_slide", "layout", "notes", "change"]}
BATCH_SCHEMA = {"type": "object", "properties": {
    "title": {"type": "string"}, "subtitle": {"type": "string"},
    "slides": {"type": "array", "items": SLIDE_SCHEMA}}, "required": ["slides"]}

MAP_PROMPT = """You convert a presentation someone made themselves into the evoila US house style. evoila US is a VMware by Broadcom specialist services firm in North America. You see each original slide as a picture, plus what we pulled out of the file (text, native chart data, tables, and the visuals on the slide with their position as fractions of the slide: x, y, width, height).

Your job: for every original slide, choose the evoila layout that carries the same idea and fill it, so the new deck says what the old one said, in the same order, in the evoila look. Then call submit_slides.

WHO AND WHAT
Presenter(s): {presenters}
Audience: {audience}
Company or event it is for (may be empty): {company}
Wording: {wording}
Anything the requester told us (may be empty): {notes}
{deck_line}

EVOILA LAYOUTS
- cover: the original title slide. Only the deck title and subtitle are used (put them in the top-level title and subtitle). The presenters and meeting line are added for you; drop the old logo and date.
- agenda: title + 3 to 6 items (heading only, under 40 characters).
- statement: dark slide, one big message. kicker (under 30 characters, optional), statement (under 110 characters), support (one sentence under 140 characters, optional). Use for section dividers and one-line message slides.
- bullets: title + 3 to 6 items; each item an optional heading (2 to 4 words) and text (under 110 characters).
- columns: title + 2 or 3 items; each a heading (under 32 characters) and 2 to 4 points (under 75 characters each). Use for side by side groups.
- cards: title + 3 or 4 items; each a heading (under 30 characters), text (under 95 characters) and an icon from the allowed list. Use for a set of parallel ideas, pillars or features.
- numbers: title + 2 to 4 items, each a figure (short, like 40% or 1,400) and heading (the label, under 60 characters); source if the original names one. For KPI and big-number slides.
- steps: title + 3 to 5 items, each a heading (under 28 characters) and text (under 90 characters, may be empty). For chevrons, arrows, timelines, phases, numbered processes.
- comparison: title + headers (2 or 3 column names) + 2 to 5 rows, each a label and one short cell per header. For two or three options compared on the same points.
- table: title + headers (2 to 6 column names; the first may be empty when rows have labels) + 2 to 8 rows, each a label (the first column) and the remaining cells. For any table that does not fit comparison. Copy every cell exactly.
- chart: title + chart_ref (the id of a native chart we extracted, like C1) + chart_kind (column, bar, line or pie; keep the original kind unless another reads better) + optional support (one sentence takeaway, under 120 characters) + source if the original names one. Only for charts listed under "native charts"; we redraw them from their real numbers.
- hub: title + hub_center (the central idea, under 26 characters) + 3 to 6 items around it (heading under 24 characters, optional icon). For radial, hub and spoke or "how it fits together" diagrams.
- image: title + visual (the id of one visual we listed, like V2) OR crop (x, y, width, height as fractions of the original slide; use it when the visual is not listed, for example loose shapes drawn together or any PDF page) + optional caption (under 110 characters) + optional items (up to 4 bullets, heading or text) shown beside it. Keeps an original visual exactly as it was.
- quote: dark slide; statement is the quote (their words only), support the attribution.
- closing: dark slide, usually last. title (for example Next steps or Thank you) + up to 3 items (heading only, under 80 characters). The presenters' contact cards are added for you, so drop the old contact details.
Every slide except cover has a title under 60 characters.

HOW TO CONVERT VISUALS
- Redraw in evoila style whatever can be redrawn without losing meaning: bullet lists, text boxes, shapes with text, chevrons, arrows, simple diagrams, SmartArt, pillars, KPIs, native charts and tables.
- Keep as an image (layout image) what would lose meaning if redrawn: screenshots and product UI, architecture or network drawings with many connected parts, maps, org charts, photos of real people, sites or equipment that the slide is about, customer or partner logo walls, and any chart that is only a picture (no native chart data). Prefer the visual id; crop when there is no id that covers it.
- Drop decoration: stock photos used as mood, clip art, background shapes, the old template's logos, footers, page numbers and confidentiality lines. Say in change when you dropped a photo.
- A slide that mixes text and a kept visual becomes image with items for the text.
- Choose the layout from the slide's structure, not its looks: three boxes with headings are cards or columns, not bullets.

RULES
- One output slide per original slide, in the same order, with source_slide set to the original number. You may split one overloaded slide into two (both with the same source_slide) and you may merge nothing. Skip only slides that are empty or purely decorative, and say so in the change of the next slide.
- {wording_rule}
- Never invent facts, figures, customers, results, quotes or dates. Numbers come from the original. If the original has a placeholder, keep it in [brackets].
- No em dashes or en dashes; use commas, colons or periods. Write the company name evoila in lowercase.
- notes: start with the original speaker notes word for word if there are any, then a sentence on how to land the slide. If you moved or dropped anything, say what in one short sentence.
- change: one short line for the build report saying what the original slide became, for example "Chevrons -> steps (5)" or "Screenshot kept as image V1 with 2 bullets" or "Stock photo dropped".
- title and subtitle at top level: only when this batch contains the original first slide. Title is the original deck title; subtitle under 70 characters, from the original subtitle if there is one, without dates or company names.

THE SLIDES
"""

WORDING = {
    "keep": ("Keep their wording", "Keep the requester's words. Only shorten a line when it does not fit the layout's limit, and only fix the house rules below."),
    "tighten": ("Tighten the copy", "Tighten and organise the requester's words into short, plain, confident lines. Keep their meaning, facts and order."),
}


def _image_block(path, max_w=1400):
    from PIL import Image
    im = Image.open(path).convert("RGB")
    if im.width > max_w:
        im = im.resize((max_w, int(im.height * max_w / im.width)), Image.LANCZOS)
    b = io.BytesIO()
    im.save(b, "JPEG", quality=80)
    return {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                        "data": base64.b64encode(b.getvalue()).decode()}}


def _slide_text(d):
    part = {"title": d["title"], "text": d["text"][:30], "notes": d["notes"]}
    if d["tables"]:
        part["tables"] = d["tables"][:3]
    if d["charts"]:
        part["native charts"] = d["charts"]
    if d["visuals"]:
        part["visuals"] = d["visuals"][:12]
    return json.dumps(part, ensure_ascii=False)[:9000]


def map_batch(batch, pages, meta, opts, first):
    wording = opts.get("wording") if opts.get("wording") in WORDING else "keep"
    deck_line = f"Deck title (from the first slide): {meta.get('deck_title')}" if not first and meta.get("deck_title") else ""
    prompt = MAP_PROMPT.format(presenters=meta.get("presenters") or "", audience=meta.get("audience") or "",
                               company=meta.get("company") or "", wording=WORDING[wording][0],
                               notes=(meta.get("notes") or "")[:3000], deck_line=deck_line,
                               wording_rule=WORDING[wording][1])
    content = [{"type": "text", "text": prompt}]
    for d in batch:
        content.append({"type": "text", "text": f"\nORIGINAL SLIDE {d['n']}\nExtracted: {_slide_text(d)}"})
        if d["n"] - 1 < len(pages):
            content.append(_image_block(pages[d["n"] - 1]))
    content.append({"type": "text", "text": f"\nConvert slides {batch[0]['n']} to {batch[-1]['n']} now and call submit_slides once."})
    tools = [{"name": "submit_slides", "description": "Submit the converted slides for this batch.", "input_schema": BATCH_SCHEMA}]
    messages = [{"role": "user", "content": content}]
    for _ in range(3):
        resp = ai_steps._call({"model": ai_steps.MODEL, "max_tokens": 16000, "tools": tools,
                               "tool_choice": {"type": "tool", "name": "submit_slides"}, "messages": messages}, timeout=300)
        out = next((b.get("input") for b in resp.get("content", []) if b.get("type") == "tool_use"), None)
        if isinstance(out, dict) and out.get("slides"):
            return CD._clean_obj(out)
        messages.append({"role": "assistant", "content": resp.get("content", [])})
        messages.append({"role": "user", "content": [{"type": "tool_result", "tool_use_id": b["id"],
                                                      "content": "No slides came through. Call submit_slides again with every slide."}
                                                     for b in resp.get("content", []) if b.get("type") == "tool_use"]
                         or "Call submit_slides with every slide."})
    raise RuntimeError("Claude returned no slides")


def fallback_slide(d):
    """No Claude for this batch: keep the slide's words as bullets so nothing is lost."""
    lines = [x for t in d["text"] for x in re.split(r"\n+", t) if x.strip()][:6]
    if d["n"] == 1:
        return {"source_slide": 1, "layout": "cover", "notes": d["notes"], "change": "Cover (no AI step)"}
    return {"source_slide": d["n"], "layout": "bullets", "title": (d["title"] or (lines[0] if lines else "Slide"))[:60],
            "items": [{"text": re.sub(r"^\[[^\]]+\]\s*", "", x)[:110]} for x in lines[1 if not d["title"] else 0:]],
            "notes": d["notes"], "change": "Kept as bullets (the AI step failed for this slide)"}


# ---------------------------------------------------------------- 3. tidy the result into a spec custom_deck renders

def crop_visual(page_png, box, out):
    from PIL import Image
    im = Image.open(page_png).convert("RGB")
    W, H = im.size
    x, y, w, h = [max(0.0, float(v)) for v in (list(box) + [0, 0, 0, 0])[:4]]
    if w <= 0.02 or h <= 0.02:
        return None
    pad = 0.006
    l, t = max(0, int((x - pad) * W)), max(0, int((y - pad) * H))
    r, b = min(W, int((x + w + pad) * W)), min(H, int((y + h + pad) * H))
    if r - l < 40 or b - t < 40:
        return None
    im.crop((l, t, r, b)).save(out, optimize=True)
    return out


def finish(slides, src, pages, report):
    by_n = {d["n"]: d for d in src}
    done = []
    for i, s in enumerate(slides):
        n = int(s.get("source_slide") or 0)
        d = by_n.get(n) or {"charts": [], "visuals": []}
        lay = s.get("layout")
        if lay == "chart":
            ch = next((c for c in d["charts"] if c["id"] == s.get("chart_ref")), None) or (d["charts"][0] if len(d["charts"]) == 1 else None)
            if ch and ch.get("series") and ch.get("categories"):
                s["chart"] = ch
            else:                                    # no real numbers behind it: keep the picture instead
                s["layout"] = "image"
                s["visual"] = next((v["id"] for v in d["visuals"] if "chart" in v["kind"]), s.get("visual"))
                s["change"] = (s.get("change") or "") + " (chart had no data in the file, kept as image)"
                lay = "image"
        if lay == "image":
            box = None
            v = next((v for v in d["visuals"] if v["id"] == s.get("visual")), None)
            if v:
                box = v["box"]
            elif isinstance(s.get("crop"), list) and len(s["crop"]) == 4:
                box = s["crop"]
            path = None
            if box and 0 < n <= len(pages):
                path = crop_visual(pages[n - 1], box, os.path.join(rr.WORK, f"keep_{i:02d}.png"))
            if path:
                s["image_path"] = path
            else:
                s["layout"] = "bullets" if s.get("items") else "statement"
                if s["layout"] == "statement":
                    s["statement"] = s.get("caption") or s.get("title") or ""
                s["change"] = (s.get("change") or "") + " (visual could not be located, kept the text only)"
        done.append(s)
        report.append(f"Slide {n}: {s.get('change') or s.get('layout')}")
    if not done or done[0].get("layout") != "cover":
        done = [{"layout": "cover", "notes": ""}] + done
    return done


def run():
    os.makedirs(rr.WORK, exist_ok=True)
    f = rr.get_record(rr.T_REQ, rr.REC)
    try:
        opts = json.loads(f.get(CD.F_OUTLINE) or "{}")
    except Exception:
        opts = {}
    rr.update({rr.F["rcheck"]: "Converting: reading your slides."})
    path, kind, fname = download_source(f, opts)
    pages, pdf = render_pages(path, kind)
    src = extract_pptx(path) if kind == "pptx" else extract_pdf(pdf, len(pages))
    src = [d for d in src if d["n"] <= len(pages)][:MAX_SLIDES]
    if not src:
        raise RuntimeError("The deck has no slides we could read.")
    people = []
    for pid in f.get(rr.F["presenters"], [])[:2]:
        p = rr.get_record(rr.T_PRES, pid)
        people.append({"name": p.get(rr.P["name"], ""), "title": p.get(rr.P["title"], ""), "email": p.get(rr.P["email"], "")})
    prospect = (f.get(rr.F["prospect"]) or "").strip()
    when = ""
    if f.get(rr.F["date"]):
        dd = datetime.date.fromisoformat(f[rr.F["date"]])
        when = f"{dd.strftime('%B')} {dd.day}, {dd.year}"
    meta = {"people": people, "presenters": ", ".join(f"{p['name']} ({p['title']})" for p in people),
            "audience": f.get(rr.F["audience"]) or "", "company": prospect, "notes": f.get(rr.F["notes"]) or "",
            "meeting_lines": [f"Prepared for {prospect}" if prospect else "", when],
            "deck_title": opts.get("title") or src[0]["title"] or ""}
    batches = [src[i:i + BATCH] for i in range(0, len(src), BATCH)]
    results, report, title, subtitle = {}, [], "", ""

    def one(k):
        try:
            return k, map_batch(batches[k], pages, meta, opts, k == 0), None
        except Exception as e:
            return k, {"slides": [fallback_slide(d) for d in batches[k]]}, str(e)[:200]

    with cf.ThreadPoolExecutor(max_workers=4) as ex:
        for k, out, err in ex.map(one, range(len(batches))):
            results[k] = out
            if err:
                report.append(f"AI step failed for slides {batches[k][0]['n']}-{batches[k][-1]['n']}: {err}")
            if k == 0:
                title, subtitle = out.get("title") or "", out.get("subtitle") or ""
    slides = [s for k in sorted(results) for s in results[k].get("slides", []) if isinstance(s, dict)]
    slides = finish(slides, src, pages, report)
    spec = {"title": opts.get("title") or title or meta["deck_title"] or "evoila", "subtitle": subtitle, "slides": slides}
    rr.update({rr.F["plan"]: json.dumps({k: v for k, v in spec.items()}, indent=1, default=str)[:95000]})
    tpl = rr.fetch_assets()
    stamp = f.get(rr.F["date"]) or datetime.date.today().isoformat()
    base = f"evoila_{rr.safe(spec['title'] or prospect or 'Deck')}_{stamp}"
    pptx = os.path.join(rr.WORK, base + ".pptx")
    n = CD.render(spec, tpl, pptx, meta, rr.ENGINE)
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", rr.WORK, pptx],
                   check=True, timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    out_pdf = pptx[:-5] + ".pdf"
    rr.update({rr.F["pptx"]: [], rr.F["pdf"]: []})
    rr.upload(rr.F["pptx"], pptx, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    rr.upload(rr.F["pdf"], out_pdf, "application/pdf")
    head = f"Converted {len(src)} original slides into {n} evoila slides."
    rr.update({rr.F["status"]: "Ready", rr.F["log"]: rr.RUN_URL or None, rr.F["error"]: "",
               rr.F["libver"]: "converted deck", rr.F["rcheck"]: "\n".join([head] + report)[:95000]})
    print("convert done:", len(src), "->", n, "slides")
