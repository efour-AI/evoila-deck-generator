"""Deck planning and market research with the Claude API (replaces the Airtable AI steps).

Used by run_request.py when the Deck Requests row arrives without an AI slide plan, which is what the
"Deck request: plan and build" automation now sends. Needs the ANTHROPIC_API_KEY repo secret.
Optional ANTHROPIC_MODEL overrides the model.

make_plan(fields, presenter_names) -> dict {summary, slides:[{slide, talking_points}], attendees}
make_research(fields) -> dict in the shape verify_market() in run_request.py expects (relevant, stats, ...)

Nothing here prints request data: the repo is public and the build logs are visible.
"""
import json, os, re, time, urllib.request, urllib.error

API_URL = "https://api.anthropic.com/v1/messages"
MODEL = os.environ.get("ANTHROPIC_MODEL") or "claude-sonnet-5-5"


def _call(body, timeout=300):
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY repo secret is not set")
    data = json.dumps(body).encode()
    last = None
    for attempt in range(4):                       # retry rate limits and overloads
        req = urllib.request.Request(API_URL, data=data, method="POST", headers={
            "x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.loads(r.read())
        except urllib.error.HTTPError as e:
            last = f"Claude API {e.code}: {e.read()[:500].decode('utf-8', 'ignore')}"
            if e.code in (429, 500, 502, 503, 529):
                time.sleep(10 * (attempt + 1))
                continue
            raise RuntimeError(last)
        except (urllib.error.URLError, TimeoutError) as e:
            last = f"Claude API unreachable: {e}"
            time.sleep(10 * (attempt + 1))
    raise RuntimeError(last or "Claude API failed")


def _name(v):
    """Airtable select values come back as plain strings from the REST API; lists of them are joined."""
    if isinstance(v, list):
        return ", ".join(str(x.get("name", x)) if isinstance(x, dict) else str(x) for x in v)
    if isinstance(v, dict):
        return str(v.get("name", ""))
    return "" if v is None else str(v)


# ---------------------------------------------------------------- plan

PLAN_SCHEMA = {
    "type": "object",
    "properties": {
        "summary": {"type": "string"},
        "slides": {"type": "array", "items": {"type": "object", "properties": {
            "slide": {"type": "number"},
            "talking_points": {"type": "array", "items": {"type": "string"}}},
            "required": ["slide", "talking_points"]}},
        "attendees": {"type": "array", "items": {"type": "string"}},
    },
    "required": ["summary", "slides", "attendees"],
}

PLAN_PROMPT = """You plan a sales meeting deck for evoila US, a VMware/VCF, private cloud, private AI, security and hybrid operations services firm in North America. The deck is built from a fixed 13-slide library; slide copy is fixed and approved. Your job: choose which library slides to use, in what order, and write short talking points for the presenter's speaker notes.

SLIDE LIBRARY
1 Cover. Always first.
2 Why we are different: pace, scale and quality together. The promise.
3 The proof behind it: evidence for pace, scale and quality, chosen for the audience (pipeline figures for partner and Broadcom audiences only, Broadcom Global Pinnacle Partner, certifications). Figures in brackets are not yet confirmed; tell the presenter to confirm them before quoting.
4 Who we are: specialty services firm, an accountable US team that owns the engagement, listen first, and a track record line (evoila's global VCF record and what the US team delivered in earlier roles).
5 Problems we solve: private cloud modernization, private AI, hybrid operations and security (VCF framing: VCF platform problems, including secure and recover).
6 Our approach: what the customer gets from an engagement, a phased plan.
7 Built right, cheaper to run: long-term cost and return case for a clean, right-sized build.
8 What we deliver: services by business problem (private cloud, resilience, what's next, security), or VCF services end to end when the framing is VCF.
9 What we find: what our assessments typically uncover.
10 Proof: anonymized customer examples already printed on the slide.
11 Global reach: global evoila engineering behind the US team.
12 How we work with you: the relationship model for this audience.
13 How to engage: next steps and contact cards. Always last.
Two data slides (industry data before 5, their pain and what we deliver after 8) are researched and added by the build; never list them in slides, and ignore any 'Data:' entries in the requester's picks.
Security talking points, when the meeting touches security: VCF was built security first; evoila secures the platform from the infrastructure up and the applications down; lead with the business exposure, then the services; confident, never alarmist.

REQUEST
Audience: {audience}
Framing: {framing}
Company: {company}
Presenters: {presenters}
Who will be in the room: {room}
Next step the presenter wants: {next_step}
Slides the requester picked (may be empty): {picked}
Meeting notes (may be dictated, informal): {notes}
Background material the presenter pasted (email threads, RFP text, requirements; may be empty or long): {background}

RULES
- If the requester picked slides, use exactly those slides plus 1 and 13, in library order, and only write talking points. Otherwise choose 6 to 10 slides that fit the notes, background material and audience; always include 1 and 13.
- 1 to 3 talking points per slide, each one plain sentence the presenter can say, tied to what the notes and background material say about this company and the people in the room. If nothing is relevant to a slide, give one point on how to land that slide with this audience.
- If the background material contains specific requirements or questions, point the presenter to the slide that answers each one.
- If a next step is given, slide 13 talking points must ask for exactly that next step.
- Match the framing: with Company framing do not talk about VCF lifecycle services; with VCF framing keep the focus on VMware Cloud Foundation.
- Never claim past work, customers, industries served, results, figures or savings that are not stated in the request above. For slide 10, tell the presenter which example printed on the slide to lead with and why; do not describe examples yourself.
- Never mention VCSP or partner of the year awards.
- Do not promise outcomes such as minimal downtime or guaranteed savings. Describe what evoila does, not what the customer will get.
- No em dashes. Plain, direct language, no hype.
- summary: two sentences on who the meeting is for and the angle of the deck.
- attendees: the people in the room, cleaned up, each as 'Name, Title' (for example 'Dana Lee, CIO'). Use only names written in 'Who will be in the room'; never invent one. Fix capitalization and punctuation only, and always separate name and title with a comma ("mickey lettuce-CIO" becomes "Mickey Lettuce, CIO"). At most three, most senior first. Empty list if no names were given.

Return the plan by calling the submit_plan tool."""


def make_plan(f, F, presenter_names):
    prompt = PLAN_PROMPT.format(
        audience=_name(f.get(F["audience"])), framing=_name(f.get(F["framing"])),
        company=_name(f.get(F["prospect"])), presenters=", ".join(presenter_names),
        room=_name(f.get(F["room"])), next_step=_name(f.get(F["next_step"])),
        picked=_name(f.get(F["slides"])), notes=_name(f.get(F["notes"])),
        background=_name(f.get(F["background"])))
    resp = _call({
        "model": MODEL, "max_tokens": 4000, "temperature": 0.2,
        "tools": [{"name": "submit_plan", "description": "Submit the deck plan.", "input_schema": PLAN_SCHEMA}],
        "tool_choice": {"type": "tool", "name": "submit_plan"},
        "messages": [{"role": "user", "content": prompt}],
    }, timeout=180)
    for b in resp.get("content", []):
        if b.get("type") == "tool_use" and b.get("name") == "submit_plan":
            p = b.get("input") or {}
            return {"summary": p.get("summary", ""), "slides": p.get("slides", []), "attendees": p.get("attendees", [])}
    raise RuntimeError("Claude returned no plan")


# ---------------------------------------------------------------- research

RESEARCH_SCHEMA = {
    "type": "object",
    "properties": {
        "relevant": {"type": "boolean"},
        "industry": {"type": "string"},
        "headline": {"type": "string"}, "subhead": {"type": "string"},
        "takeaway": {"type": "string"}, "talking_point": {"type": "string"},
        "delivery_headline": {"type": "string"}, "delivery_takeaway": {"type": "string"},
        "delivery_talking_point": {"type": "string"},
        "stats": {"type": "array", "items": {"type": "object", "properties": {
            "figure": {"type": "string"}, "claim": {"type": "string"}, "source": {"type": "string"},
            "published": {"type": "string"}, "year": {"type": "string"}, "period_end": {"type": "string"},
            "use": {"type": "string"}, "service": {"type": "string"}, "url": {"type": "string"}},
            "required": ["figure", "claim", "source", "published", "period_end", "use", "url"]}},
    },
    "required": ["relevant", "stats"],
}

RESEARCH_PROMPT = """You research two data slides for a sales meeting deck from evoila US, a services firm for VMware Cloud Foundation, private cloud, private AI, security, disaster and cyber recovery, and hybrid operations.
Slide A (industry data) sits before 'Problems we solve' and shows the pain this company and its industry are facing, in numbers.
Slide B (their pain, what we deliver) sits after 'What we deliver' and pairs each pain figure with the evoila service that answers it.

Today's date: {today}

MEETING
Company: {company}
Company website: {website}
Deck focus (framing): {framing}
Audience: {audience}
Who will be in the room: {room}
What the presenter knows about the meeting: {notes}
Background material (may be empty): {background}

STEPS
1. Work out the company's industry and the topic of this meeting from the notes (for example: cyber recovery in healthcare, VMware licensing cost pressure in state government, private AI in financial services). If the audience is Partner or Broadcom, the pain is their customers' pain in the market they sell to.
2. Search the web for third-party data published in the last 12 months (counting back from today's date) on that pain, specific to this industry wherever possible (for example healthcare figures for a health system, not all-industry averages). Use only these kinds of sources: government and regulators (for example HHS, CISA, FBI IC3), industry bodies and associations (for example AHA, HIMSS), analyst firms (Gartner, IDC, Forrester), well known annual studies (for example IBM Cost of a Data Breach, Verizon DBIR, Uptime Institute, Flexera), and established news or trade publications reporting on them (link the publication's article). Pages on social media (LinkedIn and similar) and on vendor or consultancy sites outside those kinds are rejected by the build. Never use vendor blogs, vendor statistics roundups or marketing pages, anything from Broadcom's competitors, or evoila or Broadcom material.
   Search with the current and previous year in your queries (for example "healthcare ransomware 2026") and look for the newest edition of each annual study; last year's edition is often too old.
3. Copy each figure exactly as it appears on the source page, and use the publication date printed on that page. Never use today's date, a crawl date or a guess as the publication date.
   Prefer pages that are openly readable: government and regulator pages, association pages, and the publisher's own press release or report page. Many news sites block automated checks, and a figure the build cannot read on its page is dropped.
4. Return 10 to 12 candidate stats from at least four different sources: 5 or 6 with use 'industry' for slide A, and 5 or 6 with use 'delivery' for slide B. The build checks every figure on its page and keeps up to three per slide; a slide with fewer than two confirmed figures is left out. Never use the same figure on both slides.

SERVICES FOR SLIDE B (use the exact name)
If the deck focus is Company: 'Private cloud' (build, upgrade and migrate; storage, networking, operations and automation), 'Resilience' (disaster recovery designed and tested), 'What's next' (modern applications and private AI on the same platform), 'Security' (platform security with vDefend, hardened open-source application images, maintenance and support for open-source code, cyber recovery).
If the deck focus is VCF: 'Lead with the assessment' (a consultative start that finds what needs attention), 'VCF core' (build, upgrade and migrate; vSAN, NSX, VCF Operations, VCF Automation, vDefend security), 'Additional services' (modern applications, application security, disaster recovery, cyber compliance, private AI).
For slide B pick pain the prospect is likely feeling that one of these services directly addresses, for example outages and slow recovery (Resilience), aging or unsupported infrastructure and migration risk (Private cloud), AI ambitions held back by where data can live (What's next), ransomware, unpatched open-source code and security exposure (Security).

RULES FOR EACH STAT
- figure: the number exactly as printed on the page, short (for example 772, 4%, $7.42 million, 1 in 3).
- claim: one plain sentence fragment that completes the figure, under 18 words, using the source's own terms (a data breach stays a data breach, not a cyber attack; a global figure is not described as an industry figure). No spin.
- source: publisher and report name, short (for example: IBM, Cost of a Data Breach Report). Leave the year out; the build adds the publication month and year.
- published: the date the source page or report was published, as YYYY-MM-DD or YYYY-MM. Only sources published in the last 12 months. Prefer the newest edition of an annual study.
- year: the year from published.
- period_end: the last month the data itself covers, as YYYY-MM (for example a survey fielded through March 2026 is 2026-03; a full-year 2025 figure is 2025-12). The data period must end within the last 12 months.
- use: 'industry' (slide A) or 'delivery' (slide B).
- service: for 'delivery' stats only, one service name from the list above, exactly as written; empty for 'industry'.
- url: the exact report, study, press release or article page where the figure appears. Never a vendor product, solution or pricing page, a search result or a homepage.
- Never invent or estimate a number. If you cannot find solid recent data, return relevant false and no stats.
- Nothing about this specific company that is not public. A public fact from the company's own website or recent news is allowed as one of the stats.

SLIDE TEXT
- headline: under 55 characters for slide A, names the pressure in the industry's terms (for example: What health systems are up against). Never alarmist.
- subhead: one sentence under 120 characters naming the pressure these numbers show. Do not restate any specific figure or finding in the subhead or takeaway; the build may drop some stats.
- takeaway: one sentence under 110 characters that turns the pain toward what the meeting is about, without naming evoila or promising outcomes.
- talking_point: one sentence telling the presenter how to open with slide A.
- delivery_headline: under 50 characters for slide B (for example: Where their pain meets what we deliver).
- delivery_takeaway: one sentence under 100 characters for slide B that ties the pain to the services. Never claim past work, customers, industries served or results, and never promise outcomes.
- delivery_talking_point: one sentence telling the presenter how to walk slide B.
- relevant: false only if no data would help this meeting.
- No em dashes. Plain, direct language, no hype.

When your research is done, call the submit_research tool once with the result."""


def make_research(f, F, today):
    prompt = RESEARCH_PROMPT.format(
        today=today, company=_name(f.get(F["prospect"])), website=_name(f.get(F["website"])),
        framing=_name(f.get(F["framing"])), audience=_name(f.get(F["audience"])),
        room=_name(f.get(F["room"])), notes=_name(f.get(F["notes"])), background=_name(f.get(F["background"])))
    tools = [
        {"type": "web_search_20250305", "name": "web_search", "max_uses": 15},
        {"name": "submit_research", "description": "Submit the finished research for both data slides.",
         "input_schema": RESEARCH_SCHEMA},
    ]
    messages = [{"role": "user", "content": prompt}]
    for _ in range(6):     # server-side search can pause a long turn; continue it
        resp = _call({"model": MODEL, "max_tokens": 8000, "temperature": 0.1, "tools": tools, "messages": messages},
                     timeout=300)
        content = resp.get("content", [])
        for b in content:
            if b.get("type") == "tool_use" and b.get("name") == "submit_research":
                return b.get("input") or {"relevant": False, "stats": []}
        stop = resp.get("stop_reason")
        messages.append({"role": "assistant", "content": content})
        if stop == "pause_turn":
            continue
        # finished without calling the tool: try JSON in the text, else ask once more for the tool call
        text = " ".join(b.get("text", "") for b in content if b.get("type") == "text")
        m = re.search(r"\{.*\}", text, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        if stop == "tool_use":
            raise RuntimeError("unexpected client tool call")
        messages.append({"role": "user", "content": "Now call submit_research with your result."})
    raise RuntimeError("Claude returned no research")
