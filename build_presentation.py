from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

INK = RGBColor(0x14, 0x1B, 0x2D)
SLATE = RGBColor(0x3A, 0x46, 0x60)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MIST = RGBColor(0xEE, 0xF1, 0xF5)
AMBER = RGBColor(0xF4, 0xA3, 0x00)
RED = RGBColor(0xD7, 0x26, 0x3D)
TEAL = RGBColor(0x1B, 0x99, 0x8B)
BLUE = RGBColor(0x2E, 0x86, 0xAB)
GRAY = RGBColor(0x6C, 0x75, 0x7D)

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Courier New"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
slide_count = 0


def set_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def apply_run(run, opts):
    font = run.font
    font.name = opts.get("font", BODY)
    font.size = Pt(opts.get("size", 14))
    font.bold = opts.get("bold", False)
    font.italic = opts.get("italic", False)
    font.color.rgb = opts.get("color", INK)


def set_bullet(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    p_pr.set("marL", str(int(Inches(0.25))))
    p_pr.set("indent", str(-int(Inches(0.25))))
    p_pr.append(p_pr.makeelement(qn("a:buChar"), {"char": "•"}))


def fill_text_frame(text_frame, lines, **defaults):
    text_frame.word_wrap = True
    text_frame.margin_left = text_frame.margin_right = 0
    text_frame.margin_top = text_frame.margin_bottom = 0
    text_frame.vertical_anchor = defaults.get("anchor", MSO_ANCHOR.TOP)
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        paragraph.alignment = defaults.get("align", PP_ALIGN.LEFT)
        if defaults.get("space_after"):
            paragraph.space_after = Pt(defaults["space_after"])
        if defaults.get("bullets"):
            set_bullet(paragraph)
        runs = [(line, {})] if isinstance(line, str) else line
        for text, overrides in runs:
            run = paragraph.add_run()
            run.text = text
            apply_run(run, {**defaults, **overrides})


def add_text_box(slide, lines, x, y, w, h, **defaults):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    fill_text_frame(box.text_frame, lines, **defaults)
    return box


def add_card(slide, x, y, w, h, fill_color):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.adjustments[0] = 0.08 / min(w, h)
    set_fill(card, fill_color)
    card.line.fill.background()
    card.shadow.inherit = False
    return card


def add_labeled_shape(slide, shape_type, label, x, y, w, h, fill_color, text_color=WHITE, size=10, font=BODY):
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if shape_type == MSO_SHAPE.ROUNDED_RECTANGLE:
        shape.adjustments[0] = 0.06 / min(w, h)
    set_fill(shape, fill_color)
    shape.line.fill.background()
    shape.shadow.inherit = False
    fill_text_frame(shape.text_frame, label, size=size, bold=True, color=text_color, font=font, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return shape


def add_chip(slide, label, x, y, w, fill_color, text_color=WHITE):
    return add_labeled_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, label, x, y, w, 0.32, fill_color, text_color)


def add_circle(slide, label, x, y, size, fill_color, text_color=WHITE):
    return add_labeled_shape(slide, MSO_SHAPE.OVAL, label, x, y, size, size, fill_color, text_color, size=16)


def add_arrow(slide, x, y, w, h):
    arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(arrow, SLATE)
    arrow.line.fill.background()
    arrow.shadow.inherit = False


def add_url_card(slide, x, y, w, h, label, shows, opens, payload=None, payload_h=0, payload_size=12,
                 shows_mono=True, note=None, fill_color=MIST, label_color=SLATE, text_color=INK):
    add_card(slide, x, y, w, h, fill_color)
    add_text_box(slide, label, x + 0.2, y + 0.15, w - 0.4, 0.3, size=12, bold=True, color=label_color)
    row_y = y + 0.6
    if payload:
        add_text_box(slide, payload, x + 0.2, y + 0.5, w - 0.4, payload_h, size=payload_size, font=MONO, color=text_color)
        row_y = y + 0.5 + payload_h + 0.1
    value_x = x + 1.2
    value_w = w - 1.4
    add_chip(slide, "SHOWS", x + 0.2, row_y, 0.85, TEAL)
    add_text_box(slide, shows, value_x, row_y, value_w, 0.32, size=13, font=MONO if shows_mono else BODY, color=text_color, anchor=MSO_ANCHOR.MIDDLE)
    add_chip(slide, "OPENS", x + 0.2, row_y + 0.45, 0.85, RED)
    add_text_box(slide, opens, value_x, row_y + 0.45, value_w, 0.32, size=13, font=MONO, color=text_color, anchor=MSO_ANCHOR.MIDDLE)
    if note:
        add_text_box(slide, note, x + 0.2, row_y + 0.9, w - 0.4, 0.35, size=12, color=label_color)


def add_table(slide, header, rows_data, x, y, col_widths, row_heights, cell_styles):
    graphic_frame = slide.shapes.add_table(len(rows_data) + 1, len(header), Inches(x), Inches(y), Inches(sum(col_widths)), Inches(sum(row_heights)))
    table = graphic_frame.table
    for i, width in enumerate(col_widths):
        table.columns[i].width = Inches(width)
    for i, height in enumerate(row_heights):
        table.rows[i].height = Inches(height)
    all_rows = [header] + rows_data
    for r, row in enumerate(all_rows):
        for c, text in enumerate(row):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            if r == 0:
                cell.fill.fore_color.rgb = INK
                style = {"bold": True, "color": WHITE}
            else:
                cell.fill.fore_color.rgb = WHITE if (r - 1) % 2 == 0 else MIST
                style = cell_styles(c, text)
            text_frame = cell.text_frame
            text_frame.word_wrap = True
            run = text_frame.paragraphs[0].add_run()
            run.text = text
            apply_run(run, {"size": 14, **style})
    return table


def new_slide(title, dark=False, footer=True):
    global slide_count
    slide_count += 1
    slide = prs.slides.add_slide(BLANK)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = INK if dark else WHITE
    if title:
        add_text_box(slide, title, 0.6, 0.4, 12.13, 0.85, size=36, bold=True, font=HEAD, color=WHITE if dark else INK, anchor=MSO_ANCHOR.MIDDLE)
    if footer:
        add_text_box(slide, "Lu et al., CCS '26", 0.6, 6.95, 4, 0.35, size=10, color=SLATE)
    if slide_count > 1:
        add_text_box(slide, str(slide_count), 12.13, 6.95, 0.6, 0.35, size=10, color=MIST if dark else SLATE, align=PP_ALIGN.RIGHT)
    return slide


def add_numbered_rows(slide, items, x, y, step, text_w, text_h, circle_size=0.55, circle_color=INK, circle_text=WHITE, text_color=INK):
    for i, (title, desc) in enumerate(items):
        row_y = y + i * step
        add_circle(slide, str(i + 1), x, row_y + 0.05, circle_size, circle_color, circle_text)
        add_text_box(slide, [[(title, {"bold": True, "size": 16})], desc], x + circle_size + 0.25, row_y, text_w, text_h, size=14, color=text_color)


def set_notes(slide, notes):
    slide.notes_slide.notes_text_frame.text = notes


# Slide 1: Title
slide = new_slide(None, dark=True, footer=False)
add_text_box(slide, "Finding and Exploiting Discrepancies in QR Code Payload Processing", 0.8, 1.6, 7.4, 2.7, size=40, bold=True, font=HEAD, color=WHITE, anchor=MSO_ANCHOR.BOTTOM)
add_text_box(slide, "Lu et al., CCS '26", 0.8, 4.55, 7.4, 0.5, size=20, color=AMBER)
set_notes(slide,
    "This paper looks at what happens after your phone reads a QR code. The app shows you a link to confirm, and the authors show that the link you are shown is not always the link that opens.")

# Slide 2: Three stages
slide = new_slide("What you see is not what you open")
stages = [
    ("PRIOR WORK", "1  Decoding", "Read the code and recover the payload text", MIST),
    ("THIS PAPER", "2  Parse, validate, display", "Show the decoded URL so the user can confirm it", AMBER),
    ("PRIOR WORK", "3  Execution", "Open the URL in a browser", MIST),
]
stage_x = [0.6, 4.87, 9.13]
for i, (tag, name, desc, fill_color) in enumerate(stages):
    x = stage_x[i]
    add_card(slide, x, 1.6, 3.6, 1.75, fill_color)
    add_chip(slide, tag, x + 0.25, 1.8, 1.3, INK if fill_color == AMBER else SLATE)
    add_text_box(slide, name, x + 0.25, 2.25, 3.1, 0.4, size=18, bold=True)
    add_text_box(slide, desc, x + 0.25, 2.7, 3.1, 0.5, size=14)
    if i < 2:
        add_arrow(slide, x + 3.7, 2.28, 0.47, 0.4)
add_text_box(slide, "Why stage 2 matters", 0.6, 3.75, 7.0, 0.4, size=20, bold=True)
add_text_box(slide, [
    "The confirmation popup is the user's last check before a site loads",
    "If the display differs from what runs, even careful users are fooled",
    "Google paid a bug bounty for the Chrome display defects found here",
], 0.6, 4.3, 7.0, 2.0, size=16, bullets=True, space_after=6)
add_card(slide, 8.0, 3.75, 4.73, 2.7, INK)
add_text_box(slide, "Skipping stage 2 already earned CVEs", 8.3, 3.95, 4.13, 0.4, size=16, bold=True, color=WHITE)
add_text_box(slide, [
    [("Brave   ", {"bold": True, "color": WHITE}), ("CVE-2023-28364", {"font": MONO, "color": AMBER})],
    [("Firefox ", {"bold": True, "color": WHITE}), ("CVE-2024-0953", {"font": MONO, "color": AMBER})],
], 8.3, 4.5, 4.13, 0.8, size=16, space_after=6)
add_text_box(slide, "Both opened scanned URLs without a confirmation step.", 8.3, 5.45, 4.13, 0.75, size=14, color=MIST)
set_notes(slide,
    "Using a QR code has three stages: decoding the code, displaying the URL for the user to confirm, and opening it. "
    "Earlier research focused on decoding and execution. This paper focuses on the display stage, where the user decides whether to trust the link. "
    "That stage matters: Brave and Firefox both received CVEs, which are publicly tracked security vulnerabilities, for skipping the confirmation step.")

# Slide 3: Background
slide = new_slide("URLs and Unicode: where parsers disagree")
segments = [
    ("https://", "Scheme", SLATE),
    ("user:pwd@", "Credentials", RED),
    ("example.com", "Host", TEAL),
    (":8080", "Port", GRAY),
    ("/path", "Path", BLUE),
    ("?id=1", "Query", SLATE),
    ("#frag", "Fragment", GRAY),
]
char_w = 0.167
segment_ws = [len(text) * char_w + 0.1 for text, _, _ in segments]
total_w = sum(segment_ws) + 0.04 * (len(segments) - 1)
segment_x = (13.333 - total_w) / 2
for (text, label, fill_color), segment_w in zip(segments, segment_ws):
    add_labeled_shape(slide, MSO_SHAPE.RECTANGLE, text, segment_x, 1.55, segment_w, 0.55, fill_color, size=20, font=MONO)
    is_credentials = label == "Credentials"
    add_text_box(slide, label, segment_x, 2.18, segment_w, 0.3, size=12, align=PP_ALIGN.CENTER, color=RED if is_credentials else SLATE, bold=is_credentials)
    segment_x += segment_w + 0.04
add_text_box(slide, "Where the host ends", 0.6, 3.0, 5.9, 0.4, size=20, bold=True)
add_card(slide, 0.6, 3.55, 5.9, 1.35, MIST)
add_text_box(slide, [
    [("Everything before @ is credentials", {"bold": True, "size": 16})],
    [("sad.apple.com:a@msdn.com", {"font": MONO, "color": RED})],
    "Looks like Apple, goes to msdn.com",
], 0.85, 3.7, 5.4, 1.1, size=14)
add_card(slide, 0.6, 5.1, 5.9, 1.35, MIST)
add_text_box(slide, [
    [("Standards disagree on backslashes", {"bold": True, "size": 16})],
    "RFC 3986: a backslash is invalid",
    "WHATWG, used by most browsers: \\ becomes /",
], 0.85, 5.25, 5.4, 1.1, size=14)
add_text_box(slide, "Unicode tricks", 6.83, 3.0, 5.9, 0.4, size=20, bold=True)
add_numbered_rows(slide, [
    ("Invisible characters", "Zero-width spaces, tabs, newlines, and NULL show nothing but still change the URL"),
    ("Homographs", "Cyrillic аррӏе.com looks like apple.com; its Punycode form xn--80ak6aa92e.com does not"),
    ("Reordering (Bidi)", "Right-to-left characters change display order, so a URL can appear to say something else"),
], 6.83, 3.55, 0.97, 5.13, 0.9)
set_notes(slide,
    "In a URL, anything before the '@' character is treated as login credentials, and the host comes after it, so a URL can start with apple.com and still go to msdn.com. "
    "RFC 3986 and WHATWG are URL parsing standards, and they disagree on details like backslashes. Most browsers follow WHATWG. "
    "On the Unicode side, three kinds of characters cause trouble: invisible characters, homographs, and Bidi reordering characters.")

# Slide 4: Discrepancy types and threat model
slide = new_slide("Three kinds of discrepancy")
types = [
    ("In-app", "What an app displays differs from what it opens"),
    ("Inter-app", "Apps open the same payload differently, so an attacker can target one group of users"),
    ("Inter-OS", "The same app or system camera behaves differently on Android and iOS"),
]
type_x = [0.6, 4.74, 8.88]
for i, (title, desc) in enumerate(types):
    x = type_x[i]
    add_card(slide, x, 1.5, 3.85, 2.2, MIST)
    add_circle(slide, str(i + 1), x + 0.25, 1.75, 0.6, INK)
    add_text_box(slide, title, x + 1.05, 1.8, 2.6, 0.5, size=20, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(slide, desc, x + 0.25, 2.55, 3.35, 1.0, size=14)
add_text_box(slide, "All three break one rule: the display, the security check, and the browser should all see the same origin.", 0.6, 3.9, 12.13, 0.45, size=16, italic=True, color=SLATE)
add_text_box(slide, "Threat model", 0.6, 4.6, 7.0, 0.4, size=20, bold=True)
add_text_box(slide, [
    "The attacker places malicious QR codes by replacing printed codes or through a malicious third-party printing service",
    "The code displays a legitimate domain like apple.com but opens the attacker's site",
    "Goal: the victim loads the attacker's site",
],0.6, 5.1, 7.0, 1.6, size=14, bullets=True, space_after=6)
add_url_card(slide, 8.0, 4.6, 4.73, 1.6, "What the victim sees vs. gets", "apple.com", "attacker.com")
set_notes(slide,
    "The authors define three kinds of discrepancy. In-app: an app displays one URL but opens another. Inter-app: different apps open the same QR code to different sites. Inter-OS: the same QR code behaves differently on Android and iOS. "
    "The attacker in the threat model can place malicious QR codes, for example by replacing printed codes or through a malicious third-party printing service. "
    "The code displays a legitimate domain like apple.com but opens the attacker's site, and the goal is to get the victim to load that site.")

# Slide 5: Methodology
slide = new_slide("126 test cases, three stages measured")
add_text_box(slide, "126", 0.6, 1.4, 4.4, 1.05, size=66, bold=True)
add_text_box(slide, "test QR code payloads in three categories", 0.6, 2.45, 4.4, 0.35, size=16, color=SLATE)
categories = [
    ("Encoding", "Overlong UTF-8, invisible characters, reordering, homographs, BOM"),
    ("Confusion", "Hostname, backslash, slash, and encoding confusion from URL-parser fuzzing"),
    ("Miscellaneous", "javascript: and file: schemes, domain auto-correction, HTTP prefix recognition"),
]
for i, (title, desc) in enumerate(categories):
    y = 3.0 + i * 1.25
    add_card(slide, 0.6, y, 4.4, 1.05, MIST)
    add_text_box(slide, [[(title, {"bold": True, "size": 16})], desc], 0.8, y + 0.12, 4.0, 0.85, size=14)
tiles = [("7", "browsers"), ("4", "camera apps"), ("2", "secure scanners"), ("2", "operating systems")]
for i, (value, label) in enumerate(tiles):
    x = 5.4 + i * 1.91
    add_card(slide, x, 1.5, 1.6, 1.05, MIST)
    add_text_box(slide, value, x, 1.58, 1.6, 0.5, size=28, bold=True, align=PP_ALIGN.CENTER)
    add_text_box(slide, label, x, 2.1, 1.6, 0.35, size=12, align=PP_ALIGN.CENTER, color=SLATE)
add_text_box(slide, "Android test framework", 5.4, 2.85, 7.33, 0.4, size=20, bold=True)
steps = [
    ("INPUT", "Virtual camera", "Feeds the QR code to the app"),
    ("STAGE 1", "Frida hooks", "Raw decoded data"),
    ("STAGE 2", "Screenshot + OCR", "What the popup shows"),
    ("STAGE 3", "mitmproxy", "Domain actually visited"),
]
for i, (tag, name, desc) in enumerate(steps):
    x = 5.4 + i * 1.91
    is_input = i == 0
    text_color = WHITE if is_input else INK
    add_card(slide, x, 3.4, 1.6, 1.6, INK if is_input else MIST)
    add_text_box(slide, tag, x + 0.1, 3.52, 1.4, 0.25, size=11, bold=True, color=AMBER if is_input else SLATE, align=PP_ALIGN.CENTER)
    add_text_box(slide, name, x + 0.1, 3.8, 1.4, 0.55, size=15, bold=True, color=text_color, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(slide, desc, x + 0.1, 4.42, 1.4, 0.45, size=12, color=text_color, align=PP_ALIGN.CENTER)
    if i < 3:
        add_arrow(slide, x + 1.645, 4.05, 0.22, 0.3)
add_text_box(slide, [
    "Manual testing for Google Lens (obfuscated code) and all iOS apps (no virtual camera, no jailbreak), so no stage 1 data",
    "Custom QR generator, since standard libraries rejected the unusual payloads",
], 5.4, 5.3, 7.33, 1.35, size=14, bullets=True, space_after=6)
set_notes(slide,
    "The authors built 126 test QR codes and scanned each one with popular browsers, default camera apps, and secure scanner apps on Android and iOS. "
    "On Android they built a semi-automated framework that captures all three stages: the data the app decoded, a screenshot of the URL it displayed, and the domain the browser actually visited. "
    "Comparing the displayed URL with the visited domain reveals each discrepancy. "
    "iOS apps and Google Lens could not be instrumented, so those were tested manually.")

# Slide 6: In-app truncation
slide = new_slide("In-app: truncation hides the real host")
cards = [
    ("Long URL: Chrome (Android)", "http://sad.apple.com:a@msdn.com", "sad.apple.com...", "msdn.com", True),
    ("Tabs: Pixel Camera", "apple.com\\x09\\x09\\x09\\x09\\x09\\x09msdn.com", "apple.com...", "apple.commsdn.com", True),
    ("Fullwidth dot: Chrome (iOS)", "http://apple.com｡.evil.com", "apple.com", "apple.com.evil.com", True),
    ("NULL byte: Chrome (iOS)", "http://phishing.com\\x00@hello.apple.com", "hello.apple.com", "phishing.com", True),
    ("Backslash before @: Chrome (iOS)", "http://apple.com\\@msdn.com", "apple.com/@msdn.com", "msdn.com", True),
    ("Credentials prefix: Trend Micro (Android)", ["mail.eco163.com: Unsafe", "apple.com:a@mail.eco163.com: Safe"], "rated \"Safe\"", "mail.eco163.com", False),
]
card_x = [0.6, 4.74, 8.88]
card_y = [1.5, 4.1]
for i, (label, payload, shows, opens, shows_mono) in enumerate(cards):
    add_url_card(slide, card_x[i % 3], card_y[i // 3], 3.85, 2.3, label, shows, opens,
                 payload=payload, payload_h=0.6, payload_size=12, shows_mono=shows_mono)
set_notes(slide,
    "Each card shows a QR payload, what the app displayed, and what it actually opened. "
    "The common thread is truncation: the popup cuts the URL short. The attacker puts a trusted domain at the start and the real host where the display gets cut off. "
    "Special characters like tabs, NULL bytes, and fullwidth dots do the same thing: the display stops at that character, but the browser keeps reading. "
    "The last card is Trend Micro's scanner. It flagged a malicious site as unsafe, but rated the same site safe once apple.com was added as fake credentials.")

# Slide 7: In-app Unicode confusion
slide = new_slide("In-app: Unicode makes fake hosts look real")
add_url_card(slide, 0.6, 1.5, 7.3, 2.4, "IDN homograph: red letters are Cyrillic", "apple.com", "xn--80ak6aa92e.com",
             payload=[[("http://", {}), ("аррӏе", {"color": RED}), (".com", {})]], payload_h=0.4, payload_size=14,
             note="Popups: Firefox, Edge, Samsung Camera, Sony Camera. Address bars: Opera, Firefox")
add_url_card(slide, 0.6, 4.2, 7.3, 2.4, "Mixed script and Bidi reordering: xn--ngbc5azd is the Arabic domain ending for .web or .network","apple.com", "xn--ggbla3j.xn--ngbc5azd",
             payload="http://www.apple.com.xn--ggbla3j.xn--ngbc5azd/۰.html", payload_h=0.4, payload_size=14,
             note="Firefox (Android), Edge (Android), Samsung Camera, Sony Camera")
add_card(slide, 8.2, 1.5, 4.53, 2.4, MIST)
add_text_box(slide, [
    [("QR readers and secure scanners", {"bold": True, "size": 16})],
    "Every tested QR reader is affected",
    "Only Kaspersky on iOS shows Punycode",
    "Kaspersky on Android and Trend Micro on both OSes fall for both attacks",
], 8.45, 1.68, 4.03, 2.1, size=14, space_after=6)
add_card(slide, 8.2, 4.2, 4.53, 2.4, INK)
add_text_box(slide, [
    [("Root cause", {"bold": True, "size": 16, "color": AMBER})],
    "Apps put the raw payload straight into a text view (TextView on Android, UITextView on iOS) without escaping it or converting it to Punycode.",
], 8.45, 4.38, 4.03, 2.1, size=14, color=WHITE, space_after=6)
set_notes(slide,
    "The top card is a homograph attack: apple.com spelled with Cyrillic letters. Firefox, Edge, and the Samsung and Sony camera apps display it as apple.com, but it is a different domain. "
    "The bottom card mixes left-to-right Latin text with right-to-left Arabic. The reordering makes the URL display as apple.com while it opens an Arabic domain. "
    "Every QR reader app they tested was affected, and only Kaspersky on iOS showed the domain in Punycode. "
    "The root cause is that apps display the raw text without converting it.")


def set_cell(cell, text, fill_color, **style):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill_color
    cell.margin_left = cell.margin_right = Inches(0.1)
    cell.margin_top = cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.text_frame.word_wrap = True
    cell.text_frame.text = text
    apply_run(cell.text_frame.paragraphs[0].runs[0], {"size": 14, **style})


# Slide 8: Inter-app
slide = new_slide("Inter-app: same code, different site")
inter_app_groups = [
    ("https://apple.com\\@msdn.com", "iOS", [
        ("apple.com", "Brave, Edge, Vivaldi"),
        ("msdn.com", "Apple Camera, Chrome, Firefox, Opera, UC Browser"),
    ]),
    ("https://apple.com\\nmsdn.com", "iOS", [
        ("apple.commsdn.com", "Brave, Edge, Vivaldi"),
        ("apple.com%0Amsdn.com", "Firefox, Safari"),
        ("Web search", "Chrome, Opera, UC Browser"),
    ]),
    ("apple.com\\rmsdn.com", "iOS", [
        ("apple.commsdn.com", "Brave, Vivaldi"),
        ("apple.com", "Apple Camera"),
        ("apple.com%0Dmsdn.com", "Firefox, UC Browser"),
        ("Web search", "Edge, Opera"),
    ]),
    ("https://evil.app\\u200Bple.com", "iOS", [
        ("evil.app", "Safari"),
        ("evil.apple.com", "Other browsers"),
    ]),
    ("https://apple.com\\x00@msdn.com", "Android", [
        ("apple.com", "Opera"),
        ("Web search", "Edge"),
        ("msdn.com", "All other browsers"),
    ]),
    ("https://apple.com:443%2F@msdn.com", "Android", [
        ("apple.com", "Pixel Camera"),
        ("msdn.com", "Other camera and scanner apps"),
    ]),
]
row_count = 1 + sum(len(outcomes) for _, _, outcomes in inter_app_groups)
col_widths = [3.6, 0.9, 2.7, 4.93]
row_heights = [0.4] + [0.28] * (row_count - 1)
table = slide.shapes.add_table(row_count, len(col_widths), Inches(0.6), Inches(1.5), Inches(sum(col_widths)), Inches(sum(row_heights))).table
for i, width in enumerate(col_widths):
    table.columns[i].width = Inches(width)
for i, height in enumerate(row_heights):
    table.rows[i].height = Inches(height)
for c, text in enumerate(["Payload", "OS", "Opens", "Apps"]):
    set_cell(table.cell(0, c), text, INK, bold=True, color=WHITE)
r = 1
for group_index, (payload, os_name, outcomes) in enumerate(inter_app_groups):
    fill_color = WHITE if group_index % 2 == 0 else MIST
    last_r = r + len(outcomes) - 1
    if len(outcomes) > 1:
        table.cell(r, 0).merge(table.cell(last_r, 0))
        table.cell(r, 1).merge(table.cell(last_r, 1))
    set_cell(table.cell(r, 0), payload, fill_color, font=MONO, size=13)
    set_cell(table.cell(r, 1), os_name, fill_color, bold=True, color=BLUE if os_name == "iOS" else TEAL)
    for destination, apps in outcomes:
        destination_style = {} if destination == "Web search" else {"font": MONO, "size": 13}
        set_cell(table.cell(r, 2), destination, fill_color, **destination_style)
        set_cell(table.cell(r, 3), apps, fill_color)
        r += 1
set_notes(slide,
    "This table shows the same payload opened in different apps. Each group of rows is one payload, with each site it opened and the apps that opened it. "
    "In the first row, Brave, Edge, and Vivaldi follow the WHATWG standard and open apple.com, while five other apps open msdn.com. "
    "This lets an attacker target users of specific apps while the code looks normal to everyone else.")

# Slide 9: Inter-OS
slide = new_slide("Inter-OS: Android and iOS disagree")
add_table(slide, ["Payload", "Android opens", "iOS opens"], [
    ["https://apple.com\\x00@msdn.com", "msdn.com (except Opera)", "apple.com"],
    ["www.citi/apple.com", "www.citi/apple.com", "www.citi"],
    ["www.citi#http://apple.com", "www.citi", "apple.com"],
], 0.6, 1.5, [3.0, 2.2, 2.0], [0.45, 0.75, 0.75, 0.75], lambda c, text: {"font": MONO, "size": 13})
add_card(slide, 0.6, 4.95, 7.2, 1.6, MIST)
add_text_box(slide, [
    [("Why they differ", {"bold": True, "size": 16})],
    "Android ignores NULL bytes, while iOS stops reading at them. iOS also picks the domain by a well-known domain ending like .com or an http prefix, even across /, ?, or #.",
], 0.85, 5.1, 6.7, 1.35, size=14, space_after=4)
add_text_box(slide, "Same app, different OS", 8.2, 1.5, 4.53, 0.4, size=20, bold=True)
apps = [
    ("Brave", "Address bar shows Punycode on Android, Unicode on iOS"),
    ("Opera", "apple.com\\xC4\\x80msdn.com opens apple.commsdn.com on Android, runs a search on iOS"),
    ("Edge", "Popup on Android, opens directly on iOS. apple｡com opens apple.com on iOS, searches on Android"),
]
for i, (name, desc) in enumerate(apps):
    y = 2.05 + i * 1.53
    add_card(slide, 8.2, y, 4.53, 1.35, MIST)
    add_text_box(slide, [[(name, {"bold": True, "size": 16})], desc], 8.4, y + 0.12, 4.13, 1.15, size=14)
set_notes(slide,
    "Here the same QR code goes to different sites on Android and iOS. "
    "For example, Android ignores a NULL byte while iOS stops reading at it, so the first payload opens msdn.com on Android and apple.com on iOS. "
    "The cause is that the two operating systems parse URLs with different built-in libraries. "
    "On the right, even the same app, like Brave, Opera, or Edge, behaves differently on the two systems.")

# Slide 10: Exploitation and root causes
slide = new_slide("Real-world impact and root causes")
add_card(slide, 0.6, 1.5, 5.6, 5.1, INK)
add_text_box(slide, "One QR code, two payees", 0.85, 1.7, 5.1, 0.4, size=20, bold=True, color=WHITE)
add_text_box(slide, "paypal.com/Alice\\rEve", 0.85, 2.3, 5.1, 0.45, size=20, font=MONO, color=AMBER)
add_chip(slide, "iOS", 0.85, 3.0, 1.1, BLUE)
add_text_box(slide, "pays Alice", 2.15, 3.0, 3.8, 0.32, size=16, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_chip(slide, "ANDROID", 0.85, 3.45, 1.1, TEAL)
add_text_box(slide, "pays AliceEve", 2.15, 3.45, 3.8, 0.32, size=16, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
add_text_box(slide, [
    "A malicious third-party printer can redirect one OS's payments to the attacker",
    "A local bank's payme.hsbc links can be built the same way",
    "Cash App, Venmo, Monzo, and others use similar link formats",
], 0.85, 4.1, 5.1, 2.35, size=14, color=WHITE, bullets=True, space_after=6)
add_numbered_rows(slide, [
    ("In-app: developers ignore IDN", "Firefox's toNormalizedURL uses Uri.parse without Punycode. Trend Micro likely misreads the credentials part."),
    ("Inter-app: apps skip WHATWG", "Brave trusts gms-vision's displayValue. Opera cuts at NULL itself. Firefox uses ZXing plus Uri.parse."),
    ("Inter-OS: system APIs differ", "Android's Uri.parse treats \\ as /. iOS URL and NSURLComponents reject such URLs."),
], 6.6, 1.5, 1.75, 5.33, 1.5)
set_notes(slide,
    "The clearest real-world risk is payment links. With a single carriage return character, the same PayPal QR code pays Alice on iOS and an account named AliceEve on Android. "
    "A malicious third-party printer could print store payment codes that send Android users' money to the attacker. Other payment apps like Venmo and Cash App use similar link formats. "
    "On the right are the root causes: apps ignore internationalized domain names, apps don't follow the WHATWG standard, and Android and iOS system libraries parse URLs differently.")

# Slide 11: Mitigation and disclosure
slide = new_slide("Fixes and vendor responses")
add_text_box(slide, "One URL from scan to open", 0.6, 1.5, 6.6, 0.4, size=20, bold=True)
add_numbered_rows(slide, [
    ("Decode", "Fixed encoding; reject invalid bytes, NULLs, and ambiguous control characters"),
    ("Parse", "A WHATWG parser produces one normalized URL"),
    ("Check and display", "Use that URL for the safety check; show the full host and Punycode per Chrome's IDN rules"),
    ("Open", "After the user confirms, open that same URL"),
], 0.6, 2.05, 1.1, 5.85, 0.95, circle_color=AMBER, circle_text=INK)
add_text_box(slide, "The QR spec does not cover URL parsing, so the authors call for a shared standard.", 0.6, 6.45, 6.6, 0.35, size=12, italic=True, color=SLATE)
add_table(slide, ["Vendor", "Response"], [
    ["Google", "Confirmed Chrome iOS bugs, paid a bounty; dismissed Pixel Camera"],
    ["Samsung", "Reclassified as feature requests"],
    ["Microsoft", "Declined"],
    ["Apple", "Called it expected behavior"],
    ["Firefox", "Marked as duplicates"],
    ["Trend Micro, Kaspersky", "No response"],
    ["Bank", "Blamed the OS APIs"],
], 7.6, 1.5, [1.65, 3.48], [0.42, 0.68, 0.5, 0.5, 0.5, 0.5, 0.6, 0.5], lambda c, text: {"bold": c == 0})
set_notes(slide,
    "The authors' fix is to process each URL once: decode it with a fixed encoding, parse it with a WHATWG parser, and use that one normalized URL for the safety check, the display, and the browser. "
    "They also recommend showing Punycode for suspicious domains, as Chrome does, and creating a standard for how QR code URLs are handled. "
    "The authors reported these issues to the vendors. Google confirmed the Chrome bugs and paid a bug bounty, but most others declined or did not respond.")

# Slide 12: Strengths and limitations
slide = new_slide("Strengths and limitations")
add_card(slide, 0.6, 1.5, 5.9, 5.1, MIST)
add_circle(slide, "+", 0.85, 1.72, 0.55, TEAL)
add_text_box(slide, "Strengths", 1.6, 1.72, 4.6, 0.55, size=20, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text_box(slide, [
    "End-to-end tests in real apps, not URL parsers in isolation",
    "All three stages captured on Android with Frida, OCR, and mitmproxy",
    "OCR mismatches checked by hand",
    "Test cases drawn from earlier attacks and bug reports",
    "Code, test cases, and QR images are public",
], 0.85, 2.5, 5.4, 3.9, size=16, bullets=True, space_after=6)
add_card(slide, 6.83, 1.5, 5.9, 5.1, MIST)
add_circle(slide, "-", 7.08, 1.72, 0.55, RED)
add_text_box(slide, "Limitations", 7.83, 1.72, 4.6, 0.55, size=20, bold=True, anchor=MSO_ANCHOR.MIDDLE)
add_text_box(slide, [
    "Limited scope: selected apps, two scanner vendors, URL payloads only, no retail or e-commerce apps",
    "Some root causes are partly guesses because the apps are closed source",
    "iOS and Google Lens tested by hand, with no stage 1 data",
    "No user study: shows displays can mislead, not how often people are fooled",
],7.08, 2.5, 5.4, 3.9, size=16, bullets=True, space_after=6)
set_notes(slide,
    "The study's strength is end-to-end testing in real apps, from scanning to the site that loads, with all code and test cases made public. "
    "As for limitations, the scope was limited to selected apps, two scanner vendors, and URL payloads, and some root causes are partly guessed because the apps are closed source. "
    "iOS testing was manual, and there is no user study, so we know these displays can mislead but not how often people fall for them.")

# Slide 13: Conclusion
slide = new_slide("Conclusion", dark=True, footer=False)
add_text_box(slide, "No standard governs how apps handle QR code URLs, so the link shown and the link opened can differ within an app, across apps, and across operating systems.", 0.6, 1.6, 6.3, 2.4, size=24, color=WHITE)
add_text_box(slide, "Fix: one normalized URL for display, safety check, and execution.", 0.6, 4.3, 6.3, 0.8, size=18, italic=True, color=AMBER)
add_text_box(slide, "Code: github.com/ZeddYu/Chimera_QRCode", 0.6, 6.2, 6.3, 0.35, size=12, font=MONO, color=MIST)
add_card(slide, 7.4, 1.6, 5.33, 4.95, SLATE)
add_text_box(slide, "Future research", 7.7, 1.8, 4.73, 0.45, size=20, bold=True, color=WHITE)
add_numbered_rows(slide, [
    ("Other delivery methods", "NFC tags, SMS link previews, Data Matrix, Aztec"),
    ("User study", "How often each trick fools people, and which displays or warnings help"),
    ("Other payload types", "Wi-Fi, contact, calendar, and binary payloads"),
], 7.7, 2.45, 1.3, 4.03, 1.15, circle_size=0.5, circle_color=AMBER, circle_text=INK, text_color=WHITE)
set_notes(slide,
    "The main takeaway is that apps handle QR code URLs inconsistently within an app, across apps, and across operating systems, and attackers can exploit those gaps. "
    "Future work could test NFC tags and SMS link previews, run a user study, and test other payload types like Wi-Fi and contact codes.")

output_path = Path(__file__).with_name("QR_Code_Payload_Discrepancies.pptx")
prs.save(output_path)
print(f"Saved {slide_count} slides to {output_path.name}")
