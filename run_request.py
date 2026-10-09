"""Build one walking deck for an Airtable 'Deck Requests' record.

Run by GitHub Actions (see .github/workflows/build-deck.yml).
Env: AIRTABLE_TOKEN (secret), RECORD_ID, optional RUN_URL.

1. Read the request row and its linked presenters from Airtable.
2. Download the master template and the engine zip (slide builders + icons) from the 'Deck Assets' table.
3. Build the PPTX with engine/build_library.py, add AI talking points to speaker notes.
4. Convert to PDF with LibreOffice.
5. Upload both files back to the row and set Status = Ready (or Error with detail).
"""
import base64, datetime, io, json, os, re, subprocess, sys, traceback, zipfile
import urllib.request

BASE = "appYVahdfC3gbtArV"
T_REQ = "tbl3bFM0jwCjANMYO"
T_PRES = "tblWQmz1MD437ybqk"
T_ASSETS = "tblVXLTTiAInLHs0E"
F = {  # Deck Requests field IDs
    "name": "fldXM03WhdelyP62D", "email": "fldsh54IoNqdFDRM4", "prospect": "fldh6PvIYkHgR7m8o",
    "date": "fldehcJBsAzui66pz", "audience": "fldG2ZAIhyN5KdFAy", "framing": "fldCEoEULzI0Px2BR",
    "slides": "fldMOj09vHxxvahs5", "notes": "fldtrr3cxVa29DnrF", "status": "fldxQZCQ8OHA7sOTa",
    "plan": "fldE9rNtby6JwC4qk", "pptx": "fldxwsdZxpMOq0NiL", "pdf": "fldNi5sd61xg5OhcL",
    "log": "fldtBHsDTgcLuNQhb", "error": "fldELZ4iyycJu7mOD", "presenters": "fldyMIZMHeWaET7sz",
}
P = {"name": "fldr4qcJSOTaspQGm", "title": "fld68D6bAxXqefbSR", "email": "fldv8VzsSsoySdCds"}
A = {"name": "fldwyRdfc2x7SRlat", "file": "fld3T1O2M6dsaCeSy"}

TOKEN = os.environ["AIRTABLE_TOKEN"]
REC = os.environ["RECORD_ID"].strip()
RUN_URL = os.environ.get("RUN_URL", "")
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, "work")
ENGINE = os.path.join(HERE, "engine")


def api(method, url, body=None):
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read() or b"{}")


def get_record(table, rid):
    return api("GET", f"https://api.airtable.com/v0/{BASE}/{table}/{rid}?returnFieldsByFieldId=true")["fields"]


def update(fields):
    api("PATCH", f"https://api.airtable.com/v0/{BASE}/{T_REQ}/{REC}",
        {"fields": fields, "typecast": True})


def upload(field, path, ctype):
    with open(path, "rb") as fh:
        b64 = base64.b64encode(fh.read()).decode()
    api("POST", f"https://content.airtable.com/v0/{BASE}/{REC}/{field}/uploadAttachment",
        {"contentType": ctype, "file": b64, "filename": os.path.basename(path)})


def download(url, path):
    with urllib.request.urlopen(url, timeout=300) as r, open(path, "wb") as fh:
        fh.write(r.read())


def fetch_assets():
    rows = api("GET", f"https://api.airtable.com/v0/{BASE}/{T_ASSETS}?returnFieldsByFieldId=true")["records"]
    by = {r["fields"].get(A["name"], ""): r["fields"].get(A["file"], []) for r in rows}
    tpl = os.path.join(WORK, "template.pptx")
    download(by["Master template"][0]["url"], tpl)
    z = os.path.join(WORK, "assets.zip")
    download(by["Engine and icons"][0]["url"], z)
    zipfile.ZipFile(z).extractall(ENGINE)   # build_library.py, pptkit.py, icons/, earth_src.jpeg
    return tpl


def slide_no(label):
    m = re.match(r"\s*(\d+)", label or "")
    return int(m.group(1)) if m else None


def parse_plan(raw):
    """AI plan JSON written by the Airtable automation. Tolerates code fences and junk."""
    if not raw:
        return {}
    raw = raw.strip()
    m = re.search(r"\{.*\}", raw, re.S)
    try:
        return json.loads(m.group(0) if m else raw)
    except Exception:
        return {}


def safe(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s or "").strip("_")[:40] or "Meeting"


def main():
    os.makedirs(WORK, exist_ok=True)
    f = get_record(T_REQ, REC)
    audience = (f.get(F["audience"]) or "Executive")
    framing = (f.get(F["framing"]) or "Company")
    prospect = (f.get(F["prospect"]) or "").strip()
    date_s = f.get(F["date"])
    plan = parse_plan(f.get(F["plan"]))

    # slide order: explicit form choice wins, then AI plan, then the full 13
    chosen = sorted(n for n in (slide_no(s) for s in f.get(F["slides"], [])) if n)
    ai_slides = [slide_no(str(s.get("slide"))) for s in plan.get("slides", []) if isinstance(s, dict)]
    ai_slides = [n for n in ai_slides if n and 1 <= n <= 13]
    slides = chosen or list(dict.fromkeys(ai_slides)) or list(range(1, 14))
    if 1 not in slides:
        slides = [1] + slides           # every deck opens on the cover
    if 13 not in slides:
        slides = slides + [13]          # and closes on How to engage
    talking = {}
    for s in plan.get("slides", []):
        if isinstance(s, dict) and slide_no(str(s.get("slide"))):
            pts = s.get("talking_points") or []
            talking[slide_no(str(s.get("slide")))] = pts if isinstance(pts, list) else [str(pts)]

    # presenters and contacts
    people = []
    for pid in f.get(F["presenters"], [])[:2]:
        p = get_record(T_PRES, pid)
        people.append({"name": p.get(P["name"], ""), "title": p.get(P["title"], ""), "email": p.get(P["email"], "")})
    when = ""
    if date_s:
        d = datetime.date.fromisoformat(date_s)
        when = f"{d.strftime('%B')} {d.day}, {d.year}"
    meeting_line = "  |  ".join(x for x in [f"Prepared for {prospect}" if prospect else "", when] if x) or "[Meeting line]"
    req = {"audience": audience, "framing": framing, "slides": slides, "meeting_line": meeting_line,
           "presenters": "  |  ".join(f"{p['name']}, {p['title']}" for p in people) or "[Presenter name, title]"}
    if people:
        req["contacts"] = people
    req_path = os.path.join(WORK, "request.json")
    json.dump(req, open(req_path, "w"))

    tpl = fetch_assets()
    stamp = (date_s or datetime.date.today().isoformat())
    base_name = f"evoila_{safe(prospect)}_{audience.replace(' ', '')}_{stamp}"
    pptx = os.path.join(WORK, base_name + ".pptx")
    subprocess.run([sys.executable, os.path.join(ENGINE, "build_library.py"), tpl, pptx, "--request", req_path], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    # meeting talking points go on top of the library notes
    from pptx import Presentation
    prs = Presentation(pptx)
    for slide, n in zip(prs.slides, slides):
        pts = [p for p in talking.get(n, []) if str(p).strip()]
        if not pts:
            continue
        tf = slide.notes_slide.notes_text_frame
        lib = tf.text
        tf.text = ("FOR THIS MEETING\n" + "\n".join(f"- {p}" for p in pts)
                   + ("\n\nLIBRARY NOTES\n" + lib if lib.strip() else ""))
    if plan.get("summary"):
        tf = prs.slides[0].notes_slide.notes_text_frame
        tf.text = "DECK PLAN\n" + plan["summary"] + "\n\n" + tf.text
    prs.save(pptx)

    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", WORK, pptx],
                   check=True, timeout=240, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    pdf = pptx[:-5] + ".pdf"

    update({F["pptx"]: [], F["pdf"]: []})   # a rebuild replaces the old files instead of adding to them
    upload(F["pptx"], pptx, "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    upload(F["pdf"], pdf, "application/pdf")
    update({F["status"]: "Ready", F["log"]: RUN_URL or None, F["error"]: ""})
    print("done:", len(slides), "slides")   # no company names in logs; the repo is public


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        print("Build failed. Details are on the Airtable row (Error detail).")   # keep logs free of request data
        try:
            update({F["status"]: "Error", F["error"]: err[-3000:], F["log"]: RUN_URL or None})
        except Exception:
            print("Could not write the error to Airtable.")
        sys.exit(1)
