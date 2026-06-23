import zipfile
import re
import io
import copy
import os
from xml.sax.saxutils import escape

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

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
    """Blue bold label line (doc type / platform label)."""
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
        '<a:buChar char="●"/>'
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


# ─── XML slide manipulation ───────────────────────────────────────────────────

def _replace_step_number(xml, step_num):
    """Replace the step badge text (sz="900" white bold in blue roundRect)."""
    return re.sub(
        r'(sz="900".*?</a:rPr><a:t>)\d{1,2}(</a:t>)',
        rf'\g<1>{step_num}\2',
        xml,
        flags=re.DOTALL
    )


def _replace_title(xml, title):
    """Replace the section title text (sz="1400")."""
    return re.sub(
        r'(sz="1400".*?</a:rPr><a:t>)[^<]*(</a:t>)',
        rf'\g<1>{escape(title)}\2',
        xml,
        flags=re.DOTALL
    )


def _replace_body(xml, doc_label, bullets):
    """Replace content of the body placeholder (ph idx=1 type=body)."""
    new_body = _build_body(doc_label, bullets)
    return re.sub(
        r'(<p:ph idx="1" type="body"/>.*?<a:lstStyle/>).*?(</p:txBody>)',
        rf'\1{new_body}\2',
        xml,
        flags=re.DOTALL
    )


def _remove_embedded_pics(xml):
    """Remove <p:pic> elements (content-specific charts from original)."""
    return re.sub(r'<p:pic>.*?</p:pic>', '', xml, flags=re.DOTALL)


def _replace_cover(xml, cover):
    """Replace key fields in the cover slide."""
    replacements = {
        "Anand Desai Law": cover["client_name"],
        "Google Ads Audit | 2024 vs 2025-2026": cover["document_title"],
        "Full Advertising Account Strategy": cover["subtitle"],
    }
    for old, new in replacements.items():
        xml = xml.replace(escape(old), escape(new))
        xml = xml.replace(old, new)

    # Replace year: find the text node "2026" and replace with new year
    xml = re.sub(
        r'(<a:t>)\d{4}(</a:t>)(?=.*?<a:t>year</a:t>)',
        rf'\g<1>{escape(cover["year"])}\2',
        xml
    )
    return xml


# ─── Slide duplication ───────────────────────────────────────────────────────

def _get_slide_count(z):
    slides = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml$', n)]
    return len(slides)


def _read_slide(z, idx):
    return z.read(f'ppt/slides/slide{idx}.xml').decode('utf-8')


def _slide_rels(z, idx):
    path = f'ppt/slides/_rels/slide{idx}.xml.rels'
    if path in z.namelist():
        return z.read(path)
    return None


# ─── Main generator ──────────────────────────────────────────────────────────

def generate_pptx(parsed: dict) -> bytes:
    sections = parsed["sections"]
    doc_label = parsed["doc_label"]
    cover = parsed["cover"]

    # Template has slides 1 (cover) + slides 2-14 (13 content slides)
    # We need: 1 cover + len(sections) content slides
    TEMPLATE_CONTENT_SLIDES = 13
    TEMPLATE_TOTAL = 14

    with zipfile.ZipFile(TEMPLATE_PATH, 'r') as src:
        # Read all files into memory
        files = {}
        for name in src.namelist():
            files[name] = src.read(name)

        # ── Cover slide ──────────────────────────────────────────────────────
        cover_xml = files['ppt/slides/slide1.xml'].decode('utf-8')
        cover_xml = _replace_cover(cover_xml, cover)
        files['ppt/slides/slide1.xml'] = cover_xml.encode('utf-8')

        # ── Content slides ───────────────────────────────────────────────────
        # Use slide2 as the base template for all content slides
        base_content_xml = files['ppt/slides/slide2.xml'].decode('utf-8')
        base_content_rels = files.get('ppt/slides/_rels/slide2.xml.rels')

        needed = len(sections)

        # Build new content slides
        new_slide_xmls = {}
        new_slide_rels = {}

        for i, section in enumerate(sections):
            slide_idx = i + 2  # slides start at 2 for content

            xml = base_content_xml
            xml = _remove_embedded_pics(xml)
            xml = _replace_step_number(xml, section["number"])
            xml = _replace_title(xml, section["title"])
            xml = _replace_body(xml, doc_label, section["bullets"])

            new_slide_xmls[slide_idx] = xml.encode('utf-8')
            if base_content_rels:
                # Remove image relationship from rels
                rels_xml = base_content_rels.decode('utf-8')
                # Remove rId references to images (keep only background/theme rels)
                rels_xml = re.sub(
                    r'<Relationship[^/]*/?>',
                    lambda m: m.group() if 'slideLayout' in m.group() or 'hyperlink' in m.group() else '',
                    rels_xml
                )
                new_slide_rels[slide_idx] = rels_xml.encode('utf-8')

        # ── Rebuild files dict ───────────────────────────────────────────────
        # Remove all old content slides (2 through 14)
        for i in range(2, TEMPLATE_TOTAL + 1):
            files.pop(f'ppt/slides/slide{i}.xml', None)
            files.pop(f'ppt/slides/_rels/slide{i}.xml.rels', None)

        # Add new content slides
        for idx, xml_bytes in new_slide_xmls.items():
            files[f'ppt/slides/slide{idx}.xml'] = xml_bytes
        for idx, rels_bytes in new_slide_rels.items():
            files[f'ppt/slides/_rels/slide{idx}.xml.rels'] = rels_bytes

        # ── Update presentation.xml slide list ───────────────────────────────
        prs_xml = files['ppt/presentation.xml'].decode('utf-8')

        # Build new sldIdLst entries
        # Find existing entries and their rId patterns
        existing_entries = re.findall(r'<p:sldId[^/]*/>', prs_xml)

        # Keep only the first entry (cover slide, rId typically rId2 or similar)
        # We need to find what rId the cover uses
        cover_entry = existing_entries[0] if existing_entries else None

        # Build entries for all slides
        # Find the max id used
        all_ids = re.findall(r'id="(\d+)"', prs_xml)
        max_id = max(int(x) for x in all_ids) if all_ids else 256

        # Find rId pattern from existing entries
        # Cover is slide1, content starts at slide2
        # We'll rebuild the sldIdLst
        new_entries = []
        if cover_entry:
            new_entries.append(cover_entry)

        # For content slides, we need to update/add entries
        # Keep original entries for slides that exist
        content_entries = existing_entries[1:]  # skip cover

        for i in range(needed):
            slide_num = i + 2
            if i < len(content_entries):
                new_entries.append(content_entries[i])
            else:
                # Add new entry based on last known pattern
                last = content_entries[-1] if content_entries else cover_entry
                last_id = int(re.search(r'id="(\d+)"', last).group(1))
                last_rid = re.search(r'r:id="([^"]+)"', last).group(1)
                # Increment rId number
                rid_num = re.search(r'\d+$', last_rid)
                new_rid_num = int(rid_num.group()) + (i - len(content_entries) + 1)
                new_rid = last_rid[:rid_num.start()] + str(new_rid_num)
                new_id = last_id + (i - len(content_entries) + 1)
                new_entries.append(f'<p:sldId id="{new_id}" r:id="{new_rid}"/>')

        new_sldidlst = '<p:sldIdLst>' + ''.join(new_entries) + '</p:sldIdLst>'
        prs_xml = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', new_sldidlst, prs_xml, flags=re.DOTALL)
        files['ppt/presentation.xml'] = prs_xml.encode('utf-8')

        # ── Update [Content_Types].xml ────────────────────────────────────────
        ct_xml = files['[Content_Types].xml'].decode('utf-8')

        # Remove old content slide entries
        ct_xml = re.sub(
            r'<Override PartName="/ppt/slides/slide[2-9]\d*\.xml"[^/]*/?>',
            '',
            ct_xml
        )

        # Add new entries
        new_ct_entries = ''.join(
            f'<Override PartName="/ppt/slides/slide{i+2}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
            for i in range(needed)
        )
        ct_xml = ct_xml.replace(
            '<Override PartName="/ppt/slides/slide1.xml"',
            new_ct_entries + '<Override PartName="/ppt/slides/slide1.xml"'
        )
        files['[Content_Types].xml'] = ct_xml.encode('utf-8')

        # ── Write output zip ─────────────────────────────────────────────────
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as out:
            for name, data in files.items():
                out.writestr(name, data)

        return buf.getvalue()
