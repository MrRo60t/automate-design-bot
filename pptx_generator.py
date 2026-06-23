import zipfile
import re
import io
import os
from xml.sax.saxutils import escape

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

# Max content slides in the template (slides 2-14)
MAX_SECTIONS = 13

FONT_TAGS = (
    '<a:latin typeface="Montserrat"/>'
    '<a:ea typeface="Montserrat"/>'
    '<a:cs typeface="Montserrat"/>'
    '<a:sym typeface="Montserrat"/>'
)


# ─── XML builders ────────────────────────────────────────────────────────────

def _run(text, bold=False, color="FFFFFF", sz="1100"):
    b = "1" if bold else "0"
    return (
        f'<a:r>'
        f'<a:rPr b="{b}" lang="en-US" sz="{sz}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'{FONT_TAGS}'
        f'</a:rPr>'
        f'<a:t>{escape(text)}</a:t>'
        f'</a:r>'
    )


def _label_para(label_text):
    return (
        '<a:p>'
        '<a:pPr indent="0" lvl="0" marL="0" rtl="0" algn="l">'
        '<a:spcBef><a:spcPts val="360"/></a:spcBef>'
        '<a:spcAft><a:spcPts val="0"/></a:spcAft>'
        '<a:buNone/>'
        '</a:pPr>'
        + _run(label_text, bold=True, color="4F8EF7")
        + f'<a:endParaRPr b="1" sz="1100">'
        f'<a:solidFill><a:srgbClr val="4F8EF7"/></a:solidFill>'
        f'{FONT_TAGS}'
        f'</a:endParaRPr>'
        '</a:p>'
    )


def _bullet_para(label, text, first=False):
    spacing = "600" if first else "750"
    return (
        '<a:p>'
        f'<a:pPr indent="-298450" lvl="0" marL="457200" rtl="0" algn="l">'
        f'<a:spcBef><a:spcPts val="{spacing}"/></a:spcBef>'
        '<a:spcAft><a:spcPts val="0"/></a:spcAft>'
        '<a:buClr><a:srgbClr val="FFFFFF"/></a:buClr>'
        '<a:buSzPts val="1100"/>'
        '<a:buFont typeface="Montserrat"/>'
        '<a:buChar char="&#x25CF;"/>'
        '</a:pPr>'
        + _run(label, bold=True, color="FFFFFF")
        + _run(" ", bold=False, color="FFFFFF")
        + _run(text, bold=False, color="94A3B8")
        + '<a:endParaRPr b="0" sz="1100">'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'{FONT_TAGS}'
        '</a:endParaRPr>'
        '</a:p>'
    )


def _build_body(doc_label, bullets):
    parts = [_label_para(doc_label)]
    for i, b in enumerate(bullets):
        parts.append(_bullet_para(b["label"], b["text"], first=(i == 0)))
    return "".join(parts)


# ─── Slide text replacement ───────────────────────────────────────────────────

def _replace_step_number(xml, step_num):
    return re.sub(
        r'(sz="900".*?</a:rPr><a:t>)\d{1,2}(</a:t>)',
        rf'\g<1>{step_num}\2',
        xml,
        flags=re.DOTALL
    )


def _replace_title(xml, title):
    return re.sub(
        r'(sz="1400".*?</a:rPr><a:t>)[^<]*(</a:t>)',
        rf'\g<1>{escape(title)}\2',
        xml,
        flags=re.DOTALL
    )


def _replace_body(xml, doc_label, bullets):
    new_body = _build_body(doc_label, bullets)
    return re.sub(
        r'(<p:ph idx="1" type="body"/>.*?<a:lstStyle/>).*?(</p:txBody>)',
        rf'\1{new_body}\2',
        xml,
        flags=re.DOTALL
    )


def _replace_cover(xml, cover):
    replacements = [
        ("Anand Desai Law", cover["client_name"]),
        ("Google Ads Audit | 2024 vs 2025-2026", cover["document_title"]),
        ("Full Advertising Account Strategy", cover["subtitle"]),
    ]
    for old, new in replacements:
        xml = xml.replace(escape(old), escape(new))
        xml = xml.replace(old, new)

    # Replace year (appears once as a standalone text node)
    xml = re.sub(
        r'(<a:t>)\d{4}(</a:t>)',
        lambda m: m.group(1) + escape(cover["year"]) + m.group(2)
        if cover["year"] not in m.group() else m.group(),
        xml,
        count=1
    )
    return xml


# ─── Main generator ──────────────────────────────────────────────────────────

def generate_pptx(parsed: dict) -> bytes:
    sections = parsed["sections"][:MAX_SECTIONS]
    doc_label = parsed["doc_label"]
    cover = parsed["cover"]

    # Read template into memory preserving all files as-is
    files = {}
    compress_type = {}
    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as src:
        for item in src.infolist():
            files[item.filename] = src.read(item.filename)
            compress_type[item.filename] = item.compress_type

    # ── Cover slide ──────────────────────────────────────────────────────────
    key = 'ppt/slides/slide1.xml'
    xml = files[key].decode('utf-8')
    xml = _replace_cover(xml, cover)
    files[key] = xml.encode('utf-8')

    # ── Content slides (2 to 14) ─────────────────────────────────────────────
    for i, section in enumerate(sections):
        slide_idx = i + 2
        key = f'ppt/slides/slide{slide_idx}.xml'
        if key not in files:
            continue
        xml = files[key].decode('utf-8')
        xml = _replace_step_number(xml, section["number"])
        xml = _replace_title(xml, section["title"])
        xml = _replace_body(xml, doc_label, section["bullets"])
        files[key] = xml.encode('utf-8')

    # ── Write output preserving original zip structure ────────────────────────
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as out:
        for name, data in files.items():
            ct = compress_type.get(name, zipfile.ZIP_DEFLATED)
            out.writestr(name, data, compress_type=ct)

    return buf.getvalue()
