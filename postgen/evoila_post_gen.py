"""evoila US personal-profile post-image generator (Glen / Mark / Andrew), unbranded track.

Copied from the project's post-images/evoila_post_gen.py (2026-09-24 rule: no evoila logo on
personal-post images; logo'd images must follow the corporate design standards instead), with the
logo code removed. LOCKED: the colour palette. FLEXIBLE: everything else; vary layout post to post.
"""
import cairosvg, html

CREAM = "#F6F0EB"
VIOLET = "#1F1534"
GREEN = "#0BD7A1"
MINT = "#0BFFBE"
BLUE = "#0075E4"
LBLUE = "#0190FF"
INDIGO = "#0D09B5"
WHITE = "#FFFFFF"
SUB_L = "#5B5470"
SUB_D = "#C9C2E0"

HEAD = "Carlito"
BODY = "Carlito"
SIZES = {"4:5": (1080, 1350), "1:1": (1080, 1080), "wide": (1200, 627)}
# sweep_b and arc_l are left out on purpose: they sit under the bottom-left footer (name and title).
VARIANTS = ["ribbon_r", "arc_r", "corner_br", "block_r"]


def esc(s):
    return html.escape(s, quote=True)


def _grad(gid, a, b, c):
    return (f'<linearGradient id="{gid}" x1="0" y1="1" x2="1" y2="0">'
            f'<stop offset="0" stop-color="{a}"/><stop offset="0.55" stop-color="{b}"/>'
            f'<stop offset="1" stop-color="{c}"/></linearGradient>')


def graphic(W, H, variant, dark, accent):
    line = WHITE if dark else VIOLET
    if variant == "ribbon_r":
        defs = _grad("g", GREEN, BLUE, INDIGO)
        p = (f"M {W+40},{H*0.46} C {W*0.74},{H*0.50} {W*0.70},{H*0.70} {W*0.80},{H*0.82} "
             f"C {W*0.88},{H*0.92} {W*0.78},{H+40} {W*0.55},{H+40} L {W+40},{H+40} Z")
        ln = (f'<path d="M {W+40},{H*0.40} C {W*0.70},{H*0.45} {W*0.66},{H*0.74} {W*0.80},{H*0.92}" '
              f'fill="none" stroke="{line}" stroke-opacity="0.28" stroke-width="3"/>')
        return f'<defs>{defs}</defs><path d="{p}" fill="url(#g)"/>{ln}'
    if variant == "arc_r":
        defs = _grad("g", BLUE, GREEN, MINT)
        p = (f"M {W+40},{H*0.40} C {W*0.76},{H*0.46} {W*0.70},{H*0.72} {W*0.80},{H*0.90} "
             f"C {W*0.86},{H+40} {W*1.0},{H+40} {W+40},{H+40} Z")
        return f'<defs>{defs}</defs><path d="{p}" fill="url(#g)"/>'
    if variant == "corner_br":
        defs = _grad("g", MINT, GREEN, BLUE)
        p = (f"M {W+40},{H*0.58} C {W*0.68},{H*0.62} {W*0.60},{H*0.84} {W*0.72},{H+40} "
             f"L {W+40},{H+40} Z")
        return f'<defs>{defs}</defs><path d="{p}" fill="url(#g)"/>'
    if variant == "block_r":
        col = INDIGO if not dark else GREEN
        return (f'<rect x="{W*0.66}" y="0" width="{W*0.34+40}" height="{H}" fill="{col}"/>'
                f'<rect x="{W*0.66}" y="{H*0.5-4}" width="60" height="8" rx="4" fill="{accent}"/>')
    return ""


def kicker(x, y, text, dark):
    col = GREEN if dark else BLUE
    return (f'<text x="{x}" y="{y}" font-family="DejaVu Sans Mono" font-weight="700" '
            f'font-size="22" letter-spacing="3" fill="{col}">{esc(text.upper())}</text>')


def build(out, fmt="4:5", dark=False, variant="ribbon_r", kick="#ExplainIT",
          head=None, head_accent=None, accent=BLUE, subhead=None, bullets=None, footer=None):
    W, H = SIZES[fmt]
    M = 84 if fmt != "wide" else 70
    bg = VIOLET if dark else CREAM
    ink = WHITE if dark else VIOLET
    sub = SUB_D if dark else SUB_L
    el = [f'<rect width="{W}" height="{H}" fill="{bg}"/>',
          f'<clipPath id="fr"><rect width="{W}" height="{H}"/></clipPath>',
          f'<g clip-path="url(#fr)">{graphic(W, H, variant, dark, accent)}</g>',
          kicker(M, M + 42, kick, dark)]
    fs = 96 if fmt == "4:5" else (84 if fmt == "1:1" else 60)
    lh = fs * 1.06
    y = M + 170
    for ln in (head or []):
        col = accent if (head_accent and ln.strip() in head_accent) else ink
        el.append(f'<text x="{M}" y="{y}" font-family="{HEAD}" font-weight="700" '
                  f'font-size="{fs}" fill="{col}">{esc(ln)}</text>')
        y += lh
    if subhead:
        y += 18
        el.append(f'<text x="{M}" y="{y}" font-family="{HEAD}" font-weight="700" '
                  f'font-size="34" fill="{accent}">{esc(subhead)}</text>')
        y += 30
    if bullets:
        y += 36
        for b in bullets:
            el.append(f'<circle cx="{M+8}" cy="{y-9}" r="7" fill="{accent}"/>')
            el.append(f'<text x="{M+34}" y="{y}" font-family="{BODY}" font-weight="600" '
                      f'font-size="34" fill="{ink}">{esc(b)}</text>')
            y += 56
    if footer:
        el.append(f'<rect x="{M}" y="{H-M-78}" width="64" height="6" rx="3" fill="{accent}"/>')
        fy = H - M - 30
        for i, fl in enumerate(footer):
            wt = "700" if i == 0 else "400"
            el.append(f'<text x="{M}" y="{fy}" font-family="{BODY}" font-weight="{wt}" '
                      f'font-size="30" fill="{sub if i else ink}">{esc(fl)}</text>')
            fy += 40
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
           f'viewBox="0 0 {W} {H}">{"".join(el)}</svg>')
    cairosvg.svg2png(bytestring=svg.encode(), write_to=out, output_width=W, output_height=H)
