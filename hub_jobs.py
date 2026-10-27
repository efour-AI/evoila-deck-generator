"""Marketing hub background jobs, run by .github/workflows/hub-jobs.yml.

MODE=index-library  Reads every Resource Library row that is shown on the hub, pulls the text out of
                    its file (Word, PDF, PowerPoint) and rewrites the Hub Knowledge table, which the
                    hub's question bar answers from. Rows in Hub Knowledge whose Source file does not
                    start with "lib:" (for example the voice guide) are left alone.
MODE=build-post     A self-serve LinkedIn post (topic + bullets from the hub). Claude writes the copy
                    and hashtags in the author's voice using the voice guide row in Hub Knowledge,
                    the image is drawn with postgen/evoila_post_gen.py (or the author's own photo is
                    cropped to 4:5), and the row goes to the author for approval (Marketing Approved,
                    Self-serve). Approving it on the review page posts it immediately.

MODE=meeting-overviews  Backfill: writes a short Meeting Overview on Call Transcripts rows from the last
                    60 days that have a transcript and no overview yet. New calls get theirs from the
                    Fireflies scenario in Make.

Public repository: never print document text, post copy or request data to the log.
"""
import base64, json, os, re, sys, textwrap, time, urllib.error, urllib.parse, urllib.request

import ai_steps

BASE = "appYVahdfC3gbtArV"
TOKEN = os.environ["AIRTABLE_TOKEN"]
MODE = os.environ.get("MODE", "")
REC = (os.environ.get("RECORD_ID") or "").strip()
WORK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "work")

T_RES, T_KB, T_POST = "tblYQvkJUWlEhtg6s", "tblgN6cSCyfuzy6Cg", "tblchftgrc7MkquoD"
R = {"name": "fld2D09t4tkerjQfw", "share": "fldSdB4ez9MhWdK9G", "file": "fld4tiUzYaV2LsOdX", "show": "fldk6x6akKDbSw3nf"}
K = {"name": "fldZUzOp8hcDOvKDq", "share": "fldwuiyH0Yr2BrmtR", "text": "fldP8vXueSmJfLVWl", "use": "fldFvBplItZr27L9p",
     "src": "fldNkir4wHX1fNdhd"}
P = {"title": "fld0JVy1STI2nhmP1", "copy": "fldNpsfupzhYrXorE", "author": "fldaKLAg5mK69CnLk", "status": "fldvo8DzHp1DUO23I",
     "date": "fldCwmyaVEgKJWQtD", "image": "fldaWaM6MEP836pLx", "bullets": "fldBmUUPNlni8BdL0", "upload": "fldwMMXx9NPzKLysB",
     "gen": "flde7F6SG3x75KRTR", "self": "fldKlGOWLtGe9x0ZG", "note": "fldkZDDMTtB0prwzJ"}
TITLES = {"Glen Tindal": "CEO and Managing Director, evoila US", "Mark Maisano": "VP Alliances & Strategy, evoila US",
          "Andrew Hald": "Lead Solution Architect, VCDX #4"}
T_CALL = "tbllSce3nYaZiZfRS"
C = {"title": "fldvTkxCvJBfBxuGp", "date": "fldKxPm29yO8wvNxA", "parts": "fldwpSTIKKRJw6CwT", "text": "fldZuwTbamTYvdQAa",
     "overview": "fldb82uxuawj7h5Ei"}
OVERVIEW_PROMPT = """Below is the transcript of an evoila US meeting. Everyone at evoila can read your overview on the internal
marketing hub, including people who were not on the call, so write it for a colleague who missed the meeting.

- Start with two to four plain sentences: what the meeting was for and what was discussed.
- Then a line "Decided:" followed by up to four lines starting with "- " for decisions actually made (omit if none).
- Then a line "Next:" followed by up to four lines starting with "- " for agreed next steps with the owner's first name (omit if none).
- Only what the transcript supports. No quotes longer than ten words, no pricing figures, no personal or HR matters,
  no opinions about individuals.
- No markdown, no em dashes, write evoila in lowercase. Under 160 words.
- If the transcript is mostly silence or not a business conversation, reply with exactly: NONE
Reply with the overview text only.

MEETING TITLE: {title}
DATE: {date}
PARTICIPANTS: {parts}

TRANSCRIPT:
{text}"""
CHUNK = 90000          # Airtable long text holds 100,000 characters


def api(method, url, body=None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"})
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                return json.loads(r.read() or b"{}")
        except urllib.error.HTTPError as e:
            if e.code in (429, 500, 502, 503) and attempt < 3:
                time.sleep(3 * (attempt + 1))
                continue
            raise RuntimeError(f"Airtable {e.code}: {e.read()[:300].decode('utf-8', 'ignore')}")


def rows(table, formula=None):
    out, offset = [], None
    while True:
        q = "?returnFieldsByFieldId=true&pageSize=100"
        if formula:
            q += "&filterByFormula=" + urllib.parse.quote(formula)
        if offset:
            q += "&offset=" + offset
        r = api("GET", f"https://api.airtable.com/v0/{BASE}/{table}{q}")
        out += r.get("records", [])
        offset = r.get("offset")
        if not offset:
            return out


def download(url, path):
    with urllib.request.urlopen(url, timeout=300) as r, open(path, "wb") as fh:
        fh.write(r.read())


# ------------------------------------------------------------------ index-library

def text_of(path):
    low = path.lower()
    if low.endswith(".docx"):
        import docx
        d = docx.Document(path)
        parts = []
        body = d.element.body
        for child in body.iterchildren():                       # paragraphs and tables in document order
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                t = "".join(n.text or "" for n in child.iter() if n.tag.endswith("}t")).strip()
                if t:
                    parts.append(t)
            elif tag == "tbl":
                for tr in child.iter():
                    if tr.tag.endswith("}tr"):
                        cells = []
                        for tc in tr:
                            if tc.tag.endswith("}tc"):
                                cells.append(" ".join("".join(n.text or "" for n in p.iter() if n.tag.endswith("}t"))
                                                      for p in tc.iter() if p.tag.endswith("}p")).strip())
                        if any(cells):
                            parts.append(" | ".join(cells))
        return "\n".join(parts)
    if low.endswith(".pdf"):
        from pypdf import PdfReader
        return "\n".join((pg.extract_text() or "") for pg in PdfReader(path).pages)
    if low.endswith(".pptx"):
        from pptx import Presentation
        out = []
        for i, s in enumerate(Presentation(path).slides, 1):
            tx = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
            out.append(f"[Slide {i}] " + " | ".join(tx))
        return "\n".join(out)
    if low.endswith((".txt", ".md", ".csv")):
        return open(path, encoding="utf-8", errors="ignore").read()
    return ""


def index_library():
    os.makedirs(WORK, exist_ok=True)
    res = rows(T_RES, "{Show on hub}")
    new = []
    for r in res:
        f = r["fields"]
        name, share = f.get(R["name"], "Untitled"), f.get(R["share"], "") or "Internal only"
        for a in f.get(R["file"], []) or []:
            fn = a.get("filename", "")
            # PowerPoint copies of a Word source and the master template add nothing for answering questions
            if fn.lower().endswith(".pptx") and ("template" in fn.lower() or "library" in name.lower()):
                continue
            path = os.path.join(WORK, re.sub(r"[^A-Za-z0-9._-]", "_", fn) or "file")
            try:
                download(a["url"], path)
                t = text_of(path)
            except Exception as e:
                print(f"skipped one file: {type(e).__name__}")
                continue
            t = re.sub(r"[ \t]+", " ", t)
            t = re.sub(r"\n\s*\n+", "\n\n", t).strip()
            if len(t) < 200:
                continue
            pieces = [t[i:i + CHUNK] for i in range(0, len(t), CHUNK)]
            for n, piece in enumerate(pieces, 1):
                label = name if len(pieces) == 1 else f"{name} (part {n} of {len(pieces)})"
                new.append({K["name"]: label, K["share"]: share, K["text"]: piece, K["use"]: True, K["src"]: f"lib:{r['id']}:{fn}"})
    old = [r["id"] for r in rows(T_KB) if str(r["fields"].get(K["src"], "")).startswith("lib:")]
    for i in range(0, len(old), 10):
        ids = "&".join("records[]=" + x for x in old[i:i + 10])
        api("DELETE", f"https://api.airtable.com/v0/{BASE}/{T_KB}?{ids}")
    for i in range(0, len(new), 10):
        api("POST", f"https://api.airtable.com/v0/{BASE}/{T_KB}", {"records": [{"fields": x} for x in new[i:i + 10]]})
    print(f"indexed {len(new)} pieces from {len(res)} resources")


# ------------------------------------------------------------------ build-post

POST_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string", "description": "Hook line used as the working title, under 80 characters."},
        "copy": {"type": "string", "description": "The full post body, plain text with line breaks, without the hashtags."},
        "hashtags": {"type": "array", "items": {"type": "string"}, "description": "3 to 5 hashtags from the author's list."},
        "image": {"type": "object", "properties": {
            "headline": {"type": "string", "description": "6 to 12 words for the image, the post's idea in short. No hashtags."},
            "accent_words": {"type": "string", "description": "The last few words of the headline to colour as the accent."},
            "kicker": {"type": "string", "description": "Short label above the headline: a series hashtag such as #ExplainIT, or 1 to 3 words."},
            "subhead": {"type": "string", "description": "Optional, under 40 characters."},
            "dark": {"type": "boolean"},
            "variant": {"type": "string", "enum": ["ribbon_r", "arc_r", "corner_br", "block_r"]},
            "accent": {"type": "string", "enum": ["blue", "green", "indigo"]}},
            "required": ["headline", "kicker", "dark", "variant", "accent"]},
        "check": {"type": "string", "description": "One line: any governance rule you had to work around, or 'clean'."},
    },
    "required": ["title", "copy", "hashtags", "image", "check"],
}

POST_PROMPT = """You are writing a LinkedIn post that {author} ({title}) will publish on his own profile. He started it himself on the evoila marketing hub with a topic and a few bullets. Write it in his voice, following the voice guide exactly, then call submit_post.

VOICE GUIDE
{guide}

HIS INPUT
Topic: {topic}
Bullets / notes:
{bullets}
{photo}
RULES
- Use only what his bullets say plus general, well-known industry context. Never invent customers, results, numbers, dates, quotes or events. If a bullet has a number, keep it exactly.
- Every governance rule in the guide applies. If a bullet asks for something the guide forbids (for example naming a customer or the partner program), leave it out and say so in "check".
- No em dashes. "evoila" lowercase. No links. Plain text.
- The image headline is short, specific and readable at a glance; it does not repeat the hook word for word.
- Vary the image: pick dark or light, a variant and an accent that suit the topic.
"""


def clean_copy(t):
    t = str(t or "").replace("—", ", ").replace("–", "-")
    t = re.sub(r"[ \t]+,", ",", t)
    t = re.sub(r"[ \t]{2,}", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    t = re.sub(r"\bEvoila\b", "evoila", t)
    return t.strip()


def draft(author, topic, bullets, has_photo, guide):
    prompt = POST_PROMPT.format(author=author, title=TITLES.get(author, ""), guide=guide, topic=topic or "(none given)",
                                bullets=bullets or "(none given)",
                                photo="He attached his own photo, which will be the image. Still fill in the image fields.\n" if has_photo else "")
    tools = [{"name": "submit_post", "description": "Submit the finished post.", "input_schema": POST_SCHEMA}]
    messages = [{"role": "user", "content": prompt}]
    for _ in range(3):
        resp = ai_steps._call({"model": ai_steps.MODEL, "max_tokens": 4000, "tools": tools, "messages": messages}, timeout=180)
        content = resp.get("content", [])
        spec = next((b.get("input") for b in content if b.get("type") == "tool_use" and b.get("name") == "submit_post"), None)
        if isinstance(spec, dict) and spec.get("copy") and spec.get("image"):
            return spec
        messages.append({"role": "assistant", "content": content})
        if resp.get("stop_reason") == "tool_use":
            messages.append({"role": "user", "content": [{"type": "tool_result", "tool_use_id": b["id"],
                                                          "content": "Missing fields. Call submit_post again with everything."}
                                                         for b in content if b.get("type") == "tool_use"]})
        else:
            messages.append({"role": "user", "content": "Call submit_post now."})
    raise RuntimeError("Claude returned no post")


def render(spec, author, out_png):
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "postgen"))
    import evoila_post_gen as G
    im = spec.get("image") or {}
    head = clean_copy(im.get("headline") or spec.get("title") or "").replace("\n", " ")
    variant = im.get("variant") if im.get("variant") in G.VARIANTS else "ribbon_r"
    lines = textwrap.wrap(head, width=12 if variant == "block_r" else 17)
    if variant == "block_r" and len(lines) > 5:          # the colour block takes the right third
        variant, lines = "ribbon_r", textwrap.wrap(head, width=17)
    lines = lines[:5] or ["evoila"]
    acc_words = clean_copy(im.get("accent_words") or "").lower()
    accent_lines = {ln for ln in lines if acc_words and ln.lower() in acc_words} or {lines[-1]}
    col = {"blue": G.BLUE, "green": G.GREEN, "indigo": G.INDIGO}.get(im.get("accent"), G.BLUE)
    dark = bool(im.get("dark"))
    if dark and col == G.INDIGO:
        col = G.GREEN                                   # indigo disappears on violet
    sub = clean_copy(im.get("subhead") or "")[:44] or None
    if len(lines) >= 5:
        sub = None
    G.build(out_png, "4:5", dark=dark, variant=variant, kick=(clean_copy(im.get("kicker")) or "#InsightIT")[:28],
            head=lines, head_accent=accent_lines, accent=col, subhead=sub,
            footer=[author, TITLES.get(author, "evoila US")])


def to_jpeg(src_png_or_photo, out_jpg, crop=False):
    from PIL import Image, ImageOps
    im = Image.open(src_png_or_photo)
    im = ImageOps.exif_transpose(im).convert("RGB")
    if crop:
        im = ImageOps.fit(im, (1080, 1350), Image.LANCZOS, centering=(0.5, 0.45))
    im.save(out_jpg, "JPEG", quality=90, optimize=True, subsampling=1)   # no EXIF, no C2PA chunks


def build_post():
    os.makedirs(WORK, exist_ok=True)
    url = f"https://api.airtable.com/v0/{BASE}/{T_POST}/{REC}"
    f = api("GET", url + "?returnFieldsByFieldId=true")["fields"]
    author = f.get(P["author"]) or ""
    author = author.get("name") if isinstance(author, dict) else author
    if author not in TITLES:
        raise RuntimeError("Self-serve posts are for Glen, Mark or Andrew")
    api("PATCH", url, {"fields": {P["note"]: "Writing your post…"}})
    guide_rows = [r for r in rows(T_KB) if str(r["fields"].get(K["name"], "")).upper().startswith("VOICE GUIDE")]
    if not guide_rows:
        raise RuntimeError("The voice guide row is missing from Hub Knowledge")
    guide = guide_rows[0]["fields"].get(K["text"], "")
    photo = (f.get(P["upload"]) or [None])[0]
    spec = draft(author, f.get(P["title"], ""), f.get(P["bullets"], ""), bool(photo), guide)
    api("PATCH", url, {"fields": {P["note"]: "Drawing the image…"}})
    jpg = os.path.join(WORK, f"post_{REC}.jpg")
    if photo:
        raw = os.path.join(WORK, "photo_in")
        download(photo["url"], raw)
        to_jpeg(raw, jpg, crop=True)
    else:
        png = os.path.join(WORK, "post.png")
        render(spec, author, png)
        to_jpeg(png, jpg)
    tags = " ".join(t if t.startswith("#") else "#" + t for t in (spec.get("hashtags") or [])[:5])
    body = clean_copy(spec["copy"]) + ("\n\n" + tags if tags else "")
    with open(jpg, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode()
    api("PATCH", url, {"fields": {P["image"]: []}})
    api("POST", f"https://content.airtable.com/v0/{BASE}/{REC}/{P['image']}/uploadAttachment",
        {"contentType": "image/jpeg", "file": b64, "filename": f"{re.sub('[^A-Za-z]', '', author)}_selfserve.jpg"})
    check = clean_copy(spec.get("check") or "")
    fields = {P["copy"]: body, P["title"]: clean_copy(spec.get("title"))[:120] or f.get(P["title"], ""),
              P["status"]: "Draft" if "[TEST]" in (f.get(P["bullets"]) or "") else "Marketing Approved", P["gen"]: False,
              P["note"]: ("Ready. " + ("" if not check or check.lower().startswith("clean") else "Note: " + check))[:250]}
    if not f.get(P["date"]):
        fields[P["date"]] = time.strftime("%Y-%m-%dT%H:%M:%S.000Z", time.gmtime())
    api("PATCH", url, {"fields": fields, "typecast": True})
    print("post drafted")


def meeting_overviews():
    recs = rows(T_CALL, "AND({Meeting Overview}='', {Full Transcript}!='', IS_AFTER({Meeting Date}, DATEADD(TODAY(), -60, 'days')))")
    print(f"{len(recs)} meetings to summarize")
    done = 0
    for r in recs:
        f = r["fields"]
        prompt = OVERVIEW_PROMPT.format(title=f.get(C["title"], ""), date=f.get(C["date"], ""),
                                        parts=str(f.get(C["parts"], ""))[:1500], text=f.get(C["text"], "")[:95000])
        try:
            resp = ai_steps._call({"model": ai_steps.MODEL, "max_tokens": 1500,
                                   "messages": [{"role": "user", "content": prompt}]}, timeout=180)
            out = "".join(b.get("text", "") for b in resp.get("content", []) if b.get("type") == "text").strip()
        except Exception as e:
            print("one meeting failed:", type(e).__name__)
            continue
        if not out or out == "NONE":
            continue
        out = out.replace("\u2014", ", ").replace("Evoila", "evoila")
        api("PATCH", f"https://api.airtable.com/v0/{BASE}/{T_CALL}/{r['id']}", {"fields": {C["overview"]: out[:5000]}})
        done += 1
    print(f"wrote {done} overviews")


if __name__ == "__main__":
    try:
        if MODE == "index-library":
            index_library()
        elif MODE == "build-post":
            build_post()
        elif MODE == "meeting-overviews":
            meeting_overviews()
        else:
            raise SystemExit(f"unknown MODE {MODE!r}")
    except Exception as e:
        print("Job failed. Details are on the Airtable row.")
        print(type(e).__name__)
        if MODE == "build-post" and REC:
            try:
                api("PATCH", f"https://api.airtable.com/v0/{BASE}/{T_POST}/{REC}",
                    {"fields": {P["note"]: ("Something went wrong drafting this post. Try again, or tell Kevin. " + str(e))[:250]}})
            except Exception:
                pass
        sys.exit(1)
