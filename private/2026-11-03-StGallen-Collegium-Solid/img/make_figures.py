#!/usr/bin/env python3
"""Draw the figures of the Solid deck as SVG files next to this script.

The figures use the colours of the HSG template and embed the Gillius ADF No2 font of the hsg-lecture
skill, so they look the same in every browser. Text widths are measured with that font (Pillow).

  python make_figures.py
"""
import base64, io
from pathlib import Path
from PIL import Image, ImageFont

HERE = Path(__file__).resolve().parent
FONTS = HERE.parents[2] / ".claude/skills/hsg-lecture/assets/fonts"
REG, BOLD = FONTS / "GilliusADFNo2-Regular.otf", FONTS / "GilliusADFNo2-Bold.otf"

GREEN, DGREEN, LGREEN = "#00802F", "#0A5F2D", "#E3F0E7"
SALMON, TEAL, SAND = "#EB6969", "#73A5AF", "#E1D7C3"
INK, GREY, LINE, PALE = "#1a1a1a", "#595959", "#BFBFBF", "#F2F2F2"
SHARE = [GREEN, TEAL, SALMON]          # colour of server 1, 2, 3 and of their shares


def b64(path):
    return base64.b64encode(Path(path).read_bytes()).decode()


STYLE = (f'<style>@font-face{{font-family:G;font-weight:400;src:url(data:font/otf;base64,{b64(REG)})}}'
         f'@font-face{{font-family:G;font-weight:700;src:url(data:font/otf;base64,{b64(BOLD)})}}'
         'text{font-family:G,"Gill Sans",sans-serif}.b{font-weight:700}</style>')

_fonts = {}


def tw(text, size, bold=False):
    """Width of a text in the figure font."""
    key = (size, bold)
    if key not in _fonts:
        _fonts[key] = ImageFont.truetype(str(BOLD if bold else REG), size * 4)
    return _fonts[key].getlength(text) / 4


def markers():
    out = ""
    for name, col in (("g", GREEN), ("k", GREY), ("t", TEAL), ("s", SALMON), ("d", DGREEN)):
        out += (f'<marker id="a{name}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" '
                f'orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="{col}"/></marker>')
    return out


def svg(name, body, w=1200, h=440, title=""):
    s = (f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
         f'viewBox="0 0 {w} {h}" width="{w}" height="{h}">'
         f'<title>{title}</title><defs>{STYLE}{markers()}</defs>{body}</svg>')
    (HERE / name).write_text(s, encoding="utf-8")
    print("wrote", name)


# ---------- primitives

def text(x, y, t, size=24, anchor="middle", fill=INK, bold=False, extra=""):
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    b = ' class="b"' if bold else ""
    return f'<text x="{x}" y="{y}" font-size="{size}" text-anchor="{anchor}" fill="{fill}"{b} {extra}>{t}</text>'


def rect(x, y, w, h, fill="none", stroke=GREY, sw=2, r=10, extra=""):
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" {extra}/>'


def line(x1, y1, x2, y2, col=GREY, sw=3, mk="k", dash=False, start=False):
    d = ' stroke-dasharray="9 7"' if dash else ""
    m = f' marker-end="url(#a{mk})"' if mk else ""
    m += f' marker-start="url(#a{mk})"' if start and mk else ""
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{col}" stroke-width="{sw}" stroke-linecap="round"{d}{m}/>'


def path(d, col=GREY, sw=3, mk="k", dash=False, fill="none"):
    da = ' stroke-dasharray="9 7"' if dash else ""
    m = f' marker-end="url(#a{mk})"' if mk else ""
    return f'<path d="{d}" fill="{fill}" stroke="{col}" stroke-width="{sw}" stroke-linecap="round" stroke-linejoin="round"{da}{m}/>'


def person(cx, cy, s, fill=GREY):
    """A person (head and shoulders) centred on cx, cy, about s high."""
    r = s * 0.2
    return (f'<circle cx="{cx}" cy="{cy - s * 0.24}" r="{r}" fill="{fill}"/>'
            f'<path d="M{cx - s * 0.36} {cy + s * 0.46} Q{cx - s * 0.36} {cy + s * 0.04} {cx} {cy + s * 0.04} '
            f'Q{cx + s * 0.36} {cy + s * 0.04} {cx + s * 0.36} {cy + s * 0.46} Z" fill="{fill}"/>')


def photo(cx, cy, s, col=GREY, bg="#fff"):
    w, h = s, s * 0.78
    x, y = cx - w / 2, cy - h / 2
    return (rect(x, y, w, h, bg, col, max(2, s / 22), s * 0.08)
            + f'<path d="M{x + w * .1} {y + h * .85} L{x + w * .38} {y + h * .45} L{x + w * .58} {y + h * .7} '
              f'L{x + w * .7} {y + h * .56} L{x + w * .9} {y + h * .85} Z" fill="{col}"/>'
            + f'<circle cx="{x + w * .74}" cy="{y + h * .28}" r="{s * .08}" fill="{col}"/>')


def calendar(cx, cy, s, col=GREY, bg="#fff"):
    w, h = s * .9, s * .8
    x, y = cx - w / 2, cy - h / 2
    dots = "".join(f'<rect x="{x + w * (.17 + .24 * i)}" y="{y + h * (.45 + .25 * j)}" width="{w * .14}" height="{h * .13}" fill="{col}"/>'
                   for i in range(3) for j in range(2))
    return (rect(x, y, w, h, bg, col, max(2, s / 22), s * .06)
            + f'<rect x="{x}" y="{y}" width="{w}" height="{h * .26}" rx="{s * .06}" fill="{col}"/>' + dots)


def heart(cx, cy, s, col=GREY):
    k = s / 100
    return (f'<path transform="translate({cx - 50 * k} {cy - 46 * k}) scale({k})" '
            f'd="M50 90 C20 68 4 50 4 30 C4 14 16 4 30 4 C40 4 47 10 50 17 C53 10 60 4 70 4 C84 4 96 14 96 30 '
            f'C96 50 80 68 50 90 Z" fill="{col}"/>')


def doc(cx, cy, s, col=GREY, bg="#fff"):
    w, h = s * .74, s * .92
    x, y = cx - w / 2, cy - h / 2
    f = w * .3
    lines = "".join(f'<line x1="{x + w * .2}" y1="{y + h * t}" x2="{x + w * .8}" y2="{y + h * t}" stroke="{col}" stroke-width="{max(2, s / 20)}"/>'
                    for t in (.45, .6, .75))
    return (f'<path d="M{x} {y} H{x + w - f} L{x + w} {y + f} V{y + h} H{x} Z" fill="{bg}" stroke="{col}" stroke-width="{max(2, s / 22)}" stroke-linejoin="round"/>'
            + f'<path d="M{x + w - f} {y} V{y + f} H{x + w}" fill="none" stroke="{col}" stroke-width="{max(2, s / 22)}"/>' + lines)


def lock(cx, cy, s, col=GREY):
    w, h = s * .7, s * .52
    return (f'<path d="M{cx - s * .22} {cy} V{cy - s * .2} A{s * .22} {s * .22} 0 0 1 {cx + s * .22} {cy - s * .2} V{cy}" '
            f'fill="none" stroke="{col}" stroke-width="{s * .1}"/>'
            + f'<rect x="{cx - w / 2}" y="{cy - s * .04}" width="{w}" height="{h}" rx="{s * .08}" fill="{col}"/>'
            + f'<circle cx="{cx}" cy="{cy + h * .4}" r="{s * .07}" fill="#fff"/>')


def check(cx, cy, r, col=GREEN):
    return (f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{col}"/>'
            f'<path d="M{cx - r * .45} {cy + r * .02} L{cx - r * .12} {cy + r * .35} L{cx + r * .5} {cy - r * .35}" '
            f'fill="none" stroke="#fff" stroke-width="{r * .24}" stroke-linecap="round" stroke-linejoin="round"/>')


def cross(cx, cy, r, col=SALMON):
    a = r * .38
    return (f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{col}"/>'
            f'<path d="M{cx - a} {cy - a} L{cx + a} {cy + a} M{cx + a} {cy - a} L{cx - a} {cy + a}" '
            f'stroke="#fff" stroke-width="{r * .24}" stroke-linecap="round"/>')


def nope(cx, cy, r):
    """A quiet empty cell: grey ring."""
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{LINE}" stroke-width="3"/>'


def server(cx, cy, s, col=GREY):
    w, h = s * .9, s * .26
    out = ""
    for i in range(3):
        y = cy - s * .42 + i * (h + s * .05)
        out += (f'<rect x="{cx - w / 2}" y="{y}" width="{w}" height="{h}" rx="{s * .05}" fill="{col}"/>'
                f'<circle cx="{cx + w * .32}" cy="{y + h / 2}" r="{s * .045}" fill="#fff"/>'
                f'<line x1="{cx - w * .36}" y1="{y + h / 2}" x2="{cx}" y2="{y + h / 2}" stroke="#fff" stroke-width="{s * .04}"/>')
    return out


def browser(x, y, w, h, col=GREY, label=""):
    bar = 30
    dots = "".join(f'<circle cx="{x + 18 + i * 16}" cy="{y + bar / 2}" r="4.5" fill="#fff"/>' for i in range(3))
    lab = text(x + 70, y + bar / 2 + 7, label, 18, "start", "#fff") if label else ""
    return (rect(x, y, w, h, "#fff", col, 2.5, 10)
            + f'<path d="M{x} {y + 10} A10 10 0 0 1 {x + 10} {y} H{x + w - 10} A10 10 0 0 1 {x + w} {y + 10} V{y + bar} H{x} Z" fill="{col}"/>'
            + dots + lab)


def pod(x, y, w, h, label="Your pod", col=GREEN, fill=LGREEN, dash=False):
    d = ' stroke-dasharray="12 8"' if dash else ""
    lw = tw(label, 22, True) + 28
    return (f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="18" fill="{fill}" stroke="{col}" stroke-width="4"{d}/>'
            + f'<rect x="{x + 18}" y="{y - 17}" width="{lw}" height="34" rx="17" fill="{col}"/>'
            + text(x + 18 + lw / 2, y + 8, label, 22, fill="#fff", bold=True))


def chip(cx, cy, label, col, size=22, bold=True, fg="#fff", pad=16, h=38):
    w = tw(label, size, bold) + 2 * pad
    return (f'<rect x="{cx - w / 2}" y="{cy - h / 2}" width="{w}" height="{h}" rx="{h / 2}" fill="{col}"/>'
            + text(cx, cy + size * .36, label, size, fill=fg, bold=bold)), w


def image(x, y, w, h, path, clip=None, mime="image/jpeg"):
    c = f' clip-path="url(#{clip})"' if clip else ""
    return (f'<image x="{x}" y="{y}" width="{w}" height="{h}" preserveAspectRatio="xMidYMid slice"{c} '
            f'href="data:{mime};base64,{b64(path)}"/>')


def jpeg_bytes(path, width):
    im = Image.open(path).convert("RGB")
    im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    return buf.getvalue()


DATA = [photo, calendar, heart, doc]


def data_icons(cx, cy, s, col=GREY, which=(0, 1, 2, 3), gap=1.35):
    """Up to four data icons in a 2x2 grid centred on cx, cy."""
    out = ""
    for k, i in enumerate(which):
        dx = (k % 2 - .5) * s * gap if len(which) > 1 else 0
        dy = (k // 2 - .5) * s * gap if len(which) > 2 else 0
        out += DATA[i](cx + dx, cy + dy, s, col)
    return out


# ---------- figures

def fig_silos():
    b = person(150, 215, 150) + text(150, 335, "You", 26)
    apps = [("Social network", (0, 1)), ("Online shop", (0, 3)), ("Health app", (2, 1, 3))]
    for i, (name, which) in enumerate(apps):
        x = 450 + i * 255
        b += browser(x, 70, 215, 290, GREY)
        b += text(x + 107, 400, name, 24)
        b += data_icons(x + 107, 225, 58, SALMON, which)
        b += path(f"M150 128 V36 H{x + 107} V58", GREY, 3, "k")
    svg("fig-silos.svg", b, title="A person whose data is copied into three platforms: social network, online shop, health app")


def fig_pod():
    b = person(95, 215, 140) + text(95, 330, "You", 26)
    b += pod(205, 70, 300, 290, "Your pod")
    b += data_icons(355, 220, 70, GREEN)
    for i, name in enumerate(["Social network", "Online shop", "Health app"]):
        y = 55 + i * 120
        b += browser(830, y, 290, 95, GREY)
        b += text(975, y + 74, name, 24)
        b += line(515, 215 + (i - 1) * 60, 815, y + 55, GREEN, 3, "g", dash=True, start=True)
    svg("fig-pod.svg", b, title="Your data in your pod; three apps connect to the pod and hold no copies")


def url_row(x, y, parts, icon, size=30):
    """A web address split into coloured parts, with a label under each part."""
    out = icon
    for txt, col, label in parts:
        w = tw(txt, size) + 28
        out += rect(x, y - 30, w, 54, col if col != "none" else "#fff", col if col != "none" else LINE, 2.5, 8)
        out += text(x + w / 2, y + size * .22, txt, size, fill="#fff" if col not in ("none", LGREEN) else INK)
        out += text(x + w / 2, y + 60, label, 22, fill=GREY)
        x += w + 6
    return out


def fig_url():
    pod_url = "https://smnmyr.solidcommunity.net"
    b = url_row(130, 110, [(pod_url, GREEN, "your pod"), ("/public/assets/", LGREEN, "folders"),
                           ("publicPortrait-square.jpg", LGREEN, "file")], photo(60, 108, 80, GREY))
    b += url_row(130, 300, [(pod_url, GREEN, "your pod"), ("/profile/card#me", TEAL, "you: your WebID")],
                 person(60, 300, 95, GREY))
    svg("fig-url.svg", b, h=400, title="Two web addresses: one for a photo in the pod, one for the person (WebID)")


def fig_access():
    b = pod(70, 110, 220, 310, "Your pod")
    rows = [(photo, 175), (calendar, 270), (heart, 365)]
    cols = [("Everyone", 590, 3), ("Family", 820, 2), ("Your doctor", 1050, 1)]
    allowed = {(0, 0), (0, 1), (0, 2), (1, 1), (2, 2)}
    for ci, (name, x, n) in enumerate(cols):
        for k in range(n):
            b += person(x + (k - (n - 1) / 2) * 40, 45, 52, GREY)
        b += text(x, 108, name, 24)
    for ri, (ico, y) in enumerate(rows):
        b += ico(180, y, 62, GREEN)
        b += f'<line x1="300" y1="{y}" x2="1140" y2="{y}" stroke="{PALE}" stroke-width="2"/>'
        for ci, (_, x, _) in enumerate(cols):
            b += check(x, y, 26) if (ri, ci) in allowed else nope(x, y, 24)
    svg("fig-access.svg", b, h=440, title="A table of access rules: the photo for everyone, the calendar for family, health data for your doctor")


def fig_request():
    b = browser(40, 120, 270, 200, GREY, "Anna's app")
    b += person(175, 240, 90, GREY)
    b += pod(840, 60, 320, 330, "Your pod")
    b += photo(925, 160, 70, GREEN) + calendar(1075, 160, 66, GREEN) + heart(925, 290, 66, GREEN) + doc(1075, 290, 66, GREEN)
    # 1: request
    b += line(325, 150, 820, 150, GREY, 3.5, "k")
    c, _ = chip(470, 122, "I am Anna", GREY, 22)
    b += c
    c, _ = chip(660, 122, "portrait.jpg, please", GREY, 22)
    b += c
    # 2: check
    b += check(570, 230, 30)
    b += text(620, 239, "Anna may read it?", 24, "start", GREEN, True)
    # 3: answer
    b += line(820, 310, 325, 310, GREEN, 3.5, "g")
    b += photo(570, 288, 54, GREEN)
    for n, (x, y) in enumerate([(360, 122), (505, 230), (360, 288)], 1):
        b += f'<circle cx="{x}" cy="{y}" r="17" fill="{INK}"/>' + text(x, y + 8, str(n), 22, fill="#fff", bold=True)
    svg("fig-request.svg", b, h=420, title="Anna's app asks the pod for portrait.jpg; the pod checks that Anna may read it and sends the photo")


def cloud(cx, cy, s, col=GREY):
    k = s / 100
    return (f'<path transform="translate({cx - 50 * k} {cy - 34 * k}) scale({k})" '
            f'd="M25 68 A20 20 0 0 1 22 28 A26 26 0 0 1 72 22 A22 22 0 0 1 78 68 Z" fill="{col}"/>')


def house(cx, cy, s, col=GREY):
    k = s / 100
    return (f'<path transform="translate({cx - 50 * k} {cy - 50 * k}) scale({k})" '
            f'd="M10 48 L50 12 L90 48 L80 48 L80 90 L58 90 L58 64 L42 64 L42 90 L20 90 L20 48 Z" fill="{col}"/>')


def building(cx, cy, s, col=GREY):
    k = s / 100
    win = "".join(f'<rect x="{30 + i * 16}" y="{24 + j * 18}" width="9" height="10" fill="#fff"/>' for i in range(3) for j in range(3))
    return (f'<g transform="translate({cx - 50 * k} {cy - 50 * k}) scale({k})">'
            f'<rect x="22" y="12" width="56" height="80" fill="{col}"/>{win}<rect x="44" y="76" width="12" height="16" fill="#fff"/></g>')


def fig_hosting():
    hosts = [("A pod provider", cloud), ("Your own server", house), ("Your employer or city", building)]
    b = ""
    for i, (name, ico) in enumerate(hosts):
        x = 60 + i * 380
        b += rect(x, 150, 320, 250, PALE, PALE, 0, 14)
        b += ico(x + 160, 345, 80, GREY)
        b += text(x + 160, 428, name, 24)
        b += pod(x + 70, 170, 180, 120, "Your pod", dash=i > 0, fill=LGREEN if i == 0 else "#fff", col=GREEN if i == 0 else LINE)
        if i == 0:
            b += data_icons(x + 160, 238, 34, GREEN, gap=1.5)
    b += path("M330 205 C380 150 420 150 470 205", GREEN, 3.5, "g", dash=True)
    b += path("M710 205 C760 150 800 150 850 205", GREEN, 3.5, "g", dash=True)
    # the apps work with every host
    for i, name in enumerate(["Social network", "Online shop", "Health app"]):
        x = 225 + i * 270
        b += rect(x, 20, 210, 56, "#fff", GREY, 2.5, 10) + text(x + 105, 57, name, 22)
    svg("fig-hosting.svg", b, h=450, title="The pod can live with a pod provider, on your own server or with your employer, and can move between them; the apps stay the same")


def fig_copy():
    b = person(85, 210, 130) + text(85, 320, "You", 26)
    b += photo(200, 160, 70, GREEN)
    # e-mail envelope
    ex, ey = 300, 185
    b += line(160, 215, 290, 215, GREY, 3, "k")
    b += rect(305, 180, 100, 70, "#fff", GREY, 3, 6) + path("M305 182 L355 222 L405 182", GREY, 3, None)
    b += line(420, 215, 520, 215, GREY, 3, "k")
    # website server with copies
    b += rect(535, 70, 280, 300, PALE, PALE, 0, 14)
    b += server(600, 150, 100, GREY)
    b += text(675, 400, "Website server and its backups", 24)
    for k in range(3):
        b += photo(700 + k * 26, 245 + k * 26, 74, SALMON)
    # visitors
    for i in range(3):
        y = 75 + i * 115
        b += browser(940, y, 200, 95, GREY)
        b += photo(1040, y + 62, 46, SALMON)
        b += line(830, 220, 925, y + 50, LINE, 3, "k")
    svg("fig-copy.svg", b, h=420, title="You e-mail your photo to the website; the server and its backups keep copies that every visitor sees")


def fig_team():
    """The live team section of futuresociety.ch, with the four photos that come from pods marked."""
    raw = jpeg_bytes(HERE / "src-futuresociety-team.png", 1400)
    W, H = 1200, 627      # the crop is 2125 x 1110 px; shown 1200 wide
    k = W / 2125
    b = f'<image x="0" y="0" width="{W}" height="{H}" href="data:image/jpeg;base64,{base64.b64encode(raw).decode()}"/>'
    # photo centres in the crop (screenshot px at 2x, minus the crop offset 384, 330)
    pods = [(653 - 384, 525 - 330, "smnmyr.solidcommunity.net"), (653 - 384, 1148 - 330, "alicepal.solidcommunity.net"),
            (1178 - 384, 1148 - 330, "datapod.igrant.io"), (2234 - 384, 1148 - 330, "datapod.igrant.io")]
    for x, y, host in pods:
        cx, cy = x * k, y * k
        b += f'<circle cx="{cx}" cy="{cy}" r="{128 * k}" fill="none" stroke="{GREEN}" stroke-width="6"/>'
        c, _ = chip(cx, cy + 128 * k + 6, "pod", GREEN, 20, h=30)
        b += c
    svg("fig-team.svg", b, W, H, title="The team section of futuresociety.ch; four of the eight photos are marked as loaded from personal pods")


def fig_link():
    b = ""
    # the website, with the portrait
    b += browser(30, 40, 420, 350, GREY, "futuresociety.ch")
    b += '<clipPath id="round"><circle cx="240" cy="170" r="80"/></clipPath>'
    b += image(160, 90, 160, 160, HERE / "src-portrait.jpg", "round")
    b += text(240, 290, "Simon Mayer", 26, bold=True)
    b += f'<rect x="80" y="320" width="320" height="10" rx="5" fill="{PALE}"/><rect x="80" y="345" width="240" height="10" rx="5" fill="{PALE}"/>'
    # the link in the page
    code = ['<img src="https://smnmyr.solidcommunity.net/', '     public/assets/publicPortrait-square.jpg">']
    b += rect(500, 22, 670, 92, PALE, PALE, 0, 8)
    for n, ln in enumerate(code):
        b += (f'<text x="520" y="{60 + n * 32}" font-size="20" text-anchor="start" fill="{INK}" xml:space="preserve" '
              f'font-family="Consolas,DejaVu Sans Mono,monospace">{ln.replace("<", "&lt;").replace(">", "&gt;")}</text>')
    b += text(515, 145, "in the page: only a link", 22, "start", GREY)
    # the pod
    b += pod(760, 190, 400, 210, "smnmyr.solidcommunity.net")
    b += '<clipPath id="sq"><rect x="800" y="225" width="140" height="140" rx="10"/></clipPath>'
    b += image(800, 225, 140, 140, HERE / "src-portrait.jpg", "sq")
    b += calendar(1015, 260, 56, GREEN) + heart(1105, 260, 56, GREEN) + doc(1015, 340, 56, GREEN) + lock(1105, 345, 50, GREEN)
    # every visitor's browser fetches the photo from the pod
    b += path("M790 295 C640 300 470 230 330 185", GREEN, 4, "g")
    b += text(605, 352, "each visitor's browser", 22, fill=GREEN) + text(605, 380, "loads it from the pod", 22, fill=GREEN)
    b += path("M500 70 C470 70 440 90 420 120", LINE, 3, "k")
    svg("fig-link.svg", b, h=420, title="The website contains only a link to the photo; each visitor's browser loads the photo from Simon's pod")


def site(x, y, name, state):
    """A small website showing a portrait: 'new', 'old' or 'gone'."""
    b = browser(x, y, 160, 120, GREY)
    b += text(x + 80, y + 140 + 12, name, 20, fill=GREY)
    cx, cy = x + 80, y + 78
    if state == "gone":
        b += f'<circle cx="{cx}" cy="{cy}" r="30" fill="none" stroke="{LINE}" stroke-width="3" stroke-dasharray="6 5"/>'
    else:
        b += f'<circle cx="{cx}" cy="{cy}" r="30" fill="{GREEN if state == "new" else GREY}"/>' + person(cx, cy + 4, 46, "#fff")
    return b


def fig_control():
    b = ""
    for half, (label, state) in enumerate([("Change the photo", "new"), ("Withdraw it", "gone")]):
        ox = half * 610
        b += text(ox + 290, 36, label, 28, bold=True)
        b += pod(ox + 20, 80, 160, 150, "Your pod")
        cx, cy = ox + 100, 162
        if state == "new":
            b += f'<circle cx="{cx}" cy="{cy}" r="40" fill="{GREEN}"/>' + person(cx, cy + 6, 62, "#fff")
        else:
            b += f'<circle cx="{cx}" cy="{cy}" r="40" fill="{GREY}"/>' + person(cx, cy + 6, 62, "#fff") + lock(cx + 38, cy + 30, 44, SALMON)
        for i, name in enumerate(["futuresociety.ch", "a conference", "a news site"]):
            x, y = ox + 215 + (i % 2) * 180, 70 + (i // 2) * 175
            if i == 2:
                x = ox + 305
            b += site(x, y, name, state)
        b += line(ox + 184, 150, ox + 208, 140, GREEN if state == "new" else LINE, 3, "g" if state == "new" else "k")
    b += f'<line x1="600" y1="20" x2="600" y2="420" stroke="{LINE}" stroke-width="2"/>'
    svg("fig-control.svg", b, h=430, title="Left: a new photo in the pod appears on all three websites. Right: the photo is locked in the pod and disappears from all three")


def factory(cx, cy, s, col=GREY):
    k = s / 100
    return (f'<path transform="translate({cx - 50 * k} {cy - 50 * k}) scale({k})" '
            f'd="M8 92 V50 L32 36 V50 L56 36 V50 L80 36 V16 H92 V92 Z" fill="{col}"/>'
            + "".join(f'<rect x="{cx + (-30 + i * 22) * k}" y="{cy + 18 * k}" width="{12 * k}" height="{12 * k}" fill="#fff"/>' for i in range(3)))


def fig_chain():
    names = [("Company A", "raw material"), ("Company B", "component"), ("Company C", "product")]
    b = ""
    for i, (n, what) in enumerate(names):
        x = 150 + i * 330
        b += factory(x, 150, 140, GREY)
        b += text(x, 262, n, 26, bold=True) + text(x, 294, what, 22, fill=GREY)
        b += rect(x - 70, 325, 140, 95, PALE, PALE, 0, 12) + doc(x - 20, 372, 60, GREY) + lock(x + 32, 380, 46, SALMON)
        if i < 2:
            b += line(x + 95, 160, x + 235, 160, GREY, 4, "k")
    b += line(905, 160, 990, 160, GREY, 4, "k")
    b += f'<rect x="1005" y="110" width="110" height="100" rx="6" fill="{SAND}"/>'
    b += path("M1005 135 L1115 135 M1060 110 V135", "#fff", 4, None)
    b += text(1060, 262, "CO₂ footprint?", 26, bold=True, fill=GREEN)
    svg("fig-chain.svg", b, h=430, title="A supply chain of three companies; each keeps its data locked; the footprint of the product is unknown")


def fig_age(step):
    """Three people learn their average age; nobody tells their own. Step 1: the question; step 2: the result."""
    b = ""
    for i, name in enumerate(["Anna", "Ben", "Chiara"]):
        x = 140 + i * 220
        b += person(x, 150, 140, GREY) + text(x, 262, name, 26)
        b += rect(x - 70, 300, 140, 64, PALE, PALE, 0, 12) + text(x - 14, 344, "age", 24, fill=GREY) + lock(x + 36, 335, 40, SALMON)
    b += line(680, 200, 795, 200, LINE, 3, "k", dash=True)
    if step == 1:
        b += rect(815, 110, 330, 180, "#fff", LINE, 3, 16, 'stroke-dasharray="12 8"')
        b += text(980, 175, "Average age", 28, fill=GREY) + text(980, 250, "?", 64, fill=GREY, bold=True)
    else:
        b += rect(815, 110, 330, 180, LGREEN, GREEN, 4, 16)
        b += text(980, 175, "Average age", 28, fill=GREEN, bold=True) + text(980, 260, "41", 76, fill=GREEN, bold=True)
        b += rect(815, 320, 330, 70, GREEN, GREEN, 0, 12) + text(980, 364, "How could this be possible?", 26, fill="#fff")
    title = ("Anna, Ben and Chiara keep their ages locked; their average age is unknown" if step == 1 else
             "Anna, Ben and Chiara keep their ages locked; their average age, 41, is known. How could this be possible?")
    svg(f"fig-age-{step}.svg", b, h=420, title=title)


def fig_shares():
    xs = [560, 790, 1020]
    b = ""
    for j, x in enumerate(xs):
        b += server(x, 52, 70, SHARE[j]) + text(x, 120, f"Server {j + 1}", 24, bold=True, fill=SHARE[j])
    rows = [("Company A", 42, [17, 93, -68], 190), ("Company B", 10, [5, -20, 25], 285)]
    for name, v, sh, y in rows:
        b += factory(70, y - 4, 62, GREY) + text(118, y + 9, name, 26, "start")
        b += rect(285, y - 30, 90, 60, "#fff", GREY, 2.5, 10) + text(330, y + 12, str(v), 32, bold=True)
        b += line(385, y, 455, y, LINE, 3, "k")
        for j, x in enumerate(xs):
            c, _ = chip(x, y, f"{sh[j]}".replace("-", "−"), SHARE[j], 30, h=50, pad=26)
            b += c
    b += f'<line x1="470" y1="345" x2="1110" y2="345" stroke="{INK}" stroke-width="2"/>'
    for j, x in enumerate(xs):
        s = rows[0][2][j] + rows[1][2][j]
        b += text(x - 85, 395, "+", 30, fill=GREY) if False else ""
        c, _ = chip(x, 395, f"{s}".replace("-", "−"), SHARE[j], 30, h=50, pad=26)
        b += c
    b += text(118, 404, "Sum", 26, "start", bold=True)
    b += text(1100, 400, "= 52", 32, "start", GREEN, True) + text(1100, 432, "= 42 + 10", 20, "start", GREY)
    svg("fig-shares.svg", b, h=440, title="Company A splits 42 into 17, 93 and −68; company B splits 10 into 5, −20 and 25. Each server adds its shares: 22, 73, −43. Together they give 52")


def fig_arch():
    cx = [210, 560, 910]
    b = ""
    for i, n in enumerate(["Company A", "Company B", "Company C"]):
        b += pod(cx[i] - 140, 40, 280, 130, f"{n}'s pod")
        b += doc(cx[i] - 95, 112, 50, GREY)
        for j in range(3):
            c, _ = chip(cx[i] - 30 + j * 58, 112, f"s{j + 1}", SHARE[j], 22, h=40, pad=12)
            b += c
    for j, who in enumerate(["platform", "university", "NGO"]):
        sx = cx[j]
        for i in range(3):
            b += line(cx[i] - 30 + j * 58, 136, sx + (i - 1) * 22, 296, SHARE[j], 3, None)
        b += f'<circle cx="{sx}" cy="300" r="7" fill="{SHARE[j]}"/>'
        b += server(sx, 355, 80, SHARE[j])
        b += text(sx - 60, 350, f"Server {j + 1}", 24, "end", SHARE[j], True) + text(sx - 60, 378, who, 22, "end", GREY)
    for sx in cx[:2]:
        b += path(f"M{sx} 400 V422", GREY, 3, None)
    b += path(f"M{cx[2]} 400 V422 H1120 V352", GREY, 3, "k") + path(f"M{cx[0]} 422 H{cx[2]}", GREY, 3, None)
    b += f'<rect x="1060" y="235" width="120" height="110" rx="12" fill="{LGREEN}" stroke="{GREEN}" stroke-width="3"/>'
    b += text(1120, 282, "Product", 22, bold=True, fill=GREEN) + text(1120, 312, "footprint", 22, bold=True, fill=GREEN)
    svg("fig-arch.svg", b, h=440, title="Each company splits its data in its own pod into shares s1, s2, s3; Solid lets server 1 read only s1, server 2 only s2, server 3 only s3; together the servers compute the footprint")


if __name__ == "__main__":
    fig_age(1)
    fig_age(2)
    for f in (fig_silos, fig_pod, fig_url, fig_access, fig_request, fig_hosting, fig_copy, fig_team, fig_link,
              fig_control, fig_chain, fig_shares, fig_arch):
        f()
