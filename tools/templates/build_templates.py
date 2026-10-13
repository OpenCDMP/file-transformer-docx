"""
Builds the default DOCX templates of the transformer:

    web/src/main/resources/documents/plan.docx
    web/src/main/resources/documents/description.docx

The templates are generated from this file (styles, page setup, cover, header and footer, template
tags) so that the design is reviewable as code. Run it after changing anything below and commit
the regenerated .docx files:

    python tools/templates/build_templates.py

Requires only the Python standard library. The embedded fonts are subsets of Noto Sans (SIL Open
Font License, see fonts/OFL.txt); see README.md for how they were produced.
"""
import os
import uuid
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, '..', '..', 'web', 'src', 'main', 'resources', 'documents'))

# ---------------------------------------------------------------- design tokens

FONT = 'Noto Sans'
FONT_STRONG = 'Noto Sans SemiBold'

INK = '1F2733'        # text
MUTED = '5B6573'      # secondary text, question labels
FAINT = '8A93A0'      # header / footer, placeholders
LINE = 'DCE1E7'       # hairlines
STRIPE = 'F5F7FA'     # zebra rows
ACCENT = '2457C5'     # replaced by the configured accent colour (TemplateStyles.ACCENT_COLOR)


def tint(hex_color, factor):
    """Mixes a colour towards white, keeping `factor` of it (same formula as DocumentStyles.tint)."""
    rgb = [int(hex_color[i:i + 2], 16) for i in (0, 2, 4)]
    return ''.join('%02X' % int(round(255 - (255 - c) * factor)) for c in rgb)


ACCENT_SOFT = tint(ACCENT, 0.10)  # table headers; replaced together with the accent

PAGE_W, PAGE_H = 11906, 16838     # A4 in twips
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 1247, 1304, 1191
CONTENT_W = PAGE_W - 2 * MARGIN_X

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' \
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
DRAWING_NS = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" ' \
             'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' \
             'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' \
             'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

# ---------------------------------------------------------------- styles


def rpr(size=None, color=None, font=None, bold=False, italic=False, caps=False, spacing=None, underline=False):
    parts = []
    if font:
        parts.append('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s" w:eastAsia="%s"/>' % (font, font, font, font))
    if bold:
        parts.append('<w:b/><w:bCs/>')
    if italic:
        parts.append('<w:i/><w:iCs/>')
    if caps:
        parts.append('<w:caps/>')
    if color:
        parts.append('<w:color w:val="%s"/>' % color)
    if spacing:
        parts.append('<w:spacing w:val="%d"/>' % spacing)
    if size:
        parts.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size))
    if underline:
        parts.append('<w:u w:val="single"/>')
    return '<w:rPr>%s</w:rPr>' % ''.join(parts)


def ppr(before=None, after=None, line=None, keep_next=False, keep_lines=False, outline=None, border=None,
        indent_left=None, align=None, tabs=None, page_break_before=False, widow=True, shade=None, indent_right=None):
    parts = []
    if keep_next:
        parts.append('<w:keepNext/>')
    if keep_lines:
        parts.append('<w:keepLines/>')
    if page_break_before:
        parts.append('<w:pageBreakBefore/>')
    if border:
        parts.append('<w:pBdr>%s</w:pBdr>' % border)
    if shade:
        parts.append('<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade)
    if tabs:
        parts.append('<w:tabs>%s</w:tabs>' % tabs)
    spacing = ''
    if before is not None:
        spacing += ' w:before="%d"' % before
    if after is not None:
        spacing += ' w:after="%d"' % after
    if line is not None:
        spacing += ' w:line="%d" w:lineRule="auto"' % line
    if spacing:
        parts.append('<w:spacing%s/>' % spacing)
    if indent_left is not None or indent_right is not None:
        parts.append('<w:ind%s%s/>' % (' w:left="%d"' % indent_left if indent_left is not None else '',
                                       ' w:right="%d"' % indent_right if indent_right is not None else ''))
    if align:
        parts.append('<w:jc w:val="%s"/>' % align)
    if outline is not None:
        parts.append('<w:outlineLvl w:val="%d"/>' % outline)
    return '<w:pPr>%s</w:pPr>' % ''.join(parts)


def pstyle(style_id, name, ppr_xml='', rpr_xml='', based_on='Normal', next_style=None, hidden=False):
    return ('<w:style w:type="paragraph" w:customStyle="%d" w:styleId="%s"><w:name w:val="%s"/>%s%s<w:qFormat/>%s%s%s</w:style>'
            % (0 if style_id in BUILT_IN else 1, style_id, name,
               '<w:basedOn w:val="%s"/>' % based_on if based_on else '',
               '<w:next w:val="%s"/>' % next_style if next_style else '',
               '<w:uiPriority w:val="9"/>' if style_id.startswith('Heading') else '',
               ppr_xml, rpr_xml))


def cstyle(style_id, name, rpr_xml):
    return ('<w:style w:type="character" w:customStyle="%d" w:styleId="%s"><w:name w:val="%s"/>'
            '<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/>%s</w:style>'
            % (0 if style_id in BUILT_IN else 1, style_id, name, rpr_xml))


def border(side, color, size=4, space=0):
    return '<w:%s w:val="single" w:sz="%d" w:space="%d" w:color="%s"/>' % (side, size, space, color)


def tstyle(style_id, name, borders, header_fill, cell_margins=(70, 110, 70, 110), header_bottom=None, band=None, header_bold=True):
    top, right, bottom, left = cell_margins
    margins = ('<w:tblCellMar><w:top w:w="%d" w:type="dxa"/><w:left w:w="%d" w:type="dxa"/>'
               '<w:bottom w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/></w:tblCellMar>' % (top, left, bottom, right))
    conditional = ''
    if header_fill or header_bottom:
        tc = ''
        if header_bottom:
            tc += '<w:tcBorders>%s</w:tcBorders>' % border('bottom', header_bottom, 8)
        if header_fill:
            tc += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % header_fill
        conditional += ('<w:tblStylePr w:type="firstRow"><w:rPr>%s</w:rPr><w:tcPr>%s</w:tcPr></w:tblStylePr>'
                        % ('<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (FONT_STRONG, FONT_STRONG, FONT_STRONG) if header_bold else '', tc))
    if band:
        conditional += '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="%s"/></w:tcPr></w:tblStylePr>' % band
    return ('<w:style w:type="table" w:customStyle="1" w:styleId="%s"><w:name w:val="%s"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/>'
            '<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
            '<w:tblPr>%s<w:tblBorders>%s</w:tblBorders>%s</w:tblPr>%s</w:style>'
            % (style_id, name, '<w:tblStyleRowBandSize w:val="1"/>' if band else '', borders, margins, conditional))


BUILT_IN = {'Normal', 'Title', 'Subtitle', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6',
            'Caption', 'TOC1', 'TOC2', 'Hyperlink', 'Header', 'Footer', 'TableNormal', 'DefaultParagraphFont', 'NoList'}

TOC_TAB = '<w:tab w:val="right" w:pos="%d"/>' % CONTENT_W


def styles_xml():
    heading = lambda size, color, before, after, outline: (
        ppr(before=before, after=after, line=264, keep_next=True, keep_lines=True, outline=outline),
        rpr(size=size, color=color, font=FONT_STRONG))
    s = []
    s.append(pstyle('Normal', 'Normal', ppr(after=100, line=324), rpr(size=19, color=INK), based_on=None))
    s.append(pstyle('Title', 'Title', ppr(before=0, after=160, line=250), rpr(size=58, color=INK, font=FONT_STRONG), next_style='Subtitle'))
    s.append(pstyle('OpencdmpTitleLong', 'OpenCDMP Title Long', ppr(before=0, after=160, line=250), rpr(size=44, color=INK, font=FONT_STRONG), based_on='Title', next_style='Subtitle'))
    s.append(pstyle('Subtitle', 'Subtitle', ppr(after=0, line=300), rpr(size=21, color=MUTED)))
    # Heading 1: every blueprint section opens under a hairline
    s.append(pstyle('Heading1', 'heading 1',
                    ppr(before=600, after=200, line=264, keep_next=True, keep_lines=True, outline=0, border=border('top', LINE, 4, 18)),
                    rpr(size=34, color=INK, font=FONT_STRONG), next_style='Normal'))
    for level, (size, color, before, after) in enumerate([(26, INK, 380, 120), (22, INK, 300, 80), (20, MUTED, 260, 60), (19, MUTED, 220, 60), (19, MUTED, 200, 40)], start=2):
        p, r = heading(size, color, before, after, level - 1)
        s.append(pstyle('Heading%d' % level, 'heading %d' % level, p, r, next_style='Normal'))
    s.append(pstyle('OpencdmpKicker', 'OpenCDMP Kicker', ppr(after=140), rpr(size=15, color=ACCENT, font=FONT_STRONG, caps=True, spacing=30)))
    s.append(pstyle('OpencdmpLead', 'OpenCDMP Lead', ppr(after=160, line=348), rpr(size=21, color=INK)))
    s.append(pstyle('OpencdmpRule', 'OpenCDMP Rule',
                    '<w:pPr><w:pBdr>%s</w:pBdr><w:spacing w:before="360" w:after="360" w:line="120" w:lineRule="exact"/><w:ind w:right="%d"/></w:pPr>'
                    % (border('bottom', ACCENT, 18, 0), CONTENT_W - 900), rpr(size=4)))
    s.append(pstyle('OpencdmpCoverNote', 'OpenCDMP Cover Note', ppr(before=480, after=0), rpr(size=15, color=FAINT)))
    s.append(pstyle('OpencdmpContentsTitle', 'OpenCDMP Contents Title', ppr(after=280, keep_next=True), rpr(size=34, color=INK, font=FONT_STRONG)))
    # asset titles carry an accent bar shared with their meta line (identical borders merge)
    # identical borders make Word draw the title and its meta line as one box
    band = border('top', ACCENT_SOFT, 4, 9) + border('left', ACCENT, 30, 12) + border('bottom', ACCENT_SOFT, 4, 9) + border('right', ACCENT_SOFT, 4, 12)
    s.append(pstyle('OpencdmpAssetTitle', 'OpenCDMP Asset Title',
                    ppr(before=560, after=30, line=264, keep_next=True, keep_lines=True, outline=1, border=band, shade=ACCENT_SOFT, indent_left=260, indent_right=260),
                    rpr(size=28, color=INK, font=FONT_STRONG), next_style='OpencdmpAssetMeta'))
    s.append(pstyle('OpencdmpAssetMeta', 'OpenCDMP Asset Meta',
                    ppr(before=0, after=0, keep_next=True, border=band, shade=ACCENT_SOFT, indent_left=260, indent_right=260),
                    rpr(size=16, color=MUTED)))
    s.append(pstyle('OpencdmpAssetLead', 'OpenCDMP Asset Lead', ppr(before=240, after=120, line=324), rpr(size=19, color=INK)))
    s.append(pstyle('OpencdmpFieldSetTitle', 'OpenCDMP Field Set Title', ppr(before=280, after=80, line=276, keep_next=True, keep_lines=True), rpr(size=21, color=INK, font=FONT_STRONG)))
    s.append(pstyle('OpencdmpGuidance', 'OpenCDMP Guidance', ppr(before=20, after=120, line=276, border=border('left', ACCENT_SOFT, 18, 8), indent_left=180), rpr(size=16, color=MUTED)))
    s.append(pstyle('OpencdmpQuestion', 'OpenCDMP Question', ppr(before=120, after=30, line=264, keep_next=True), rpr(size=16, color=MUTED)))
    s.append(pstyle('OpencdmpAnswer', 'OpenCDMP Answer', ppr(after=100, line=324), rpr(size=19, color=INK)))
    s.append(pstyle('OpencdmpEntry', 'OpenCDMP Entry', ppr(before=200, after=40, keep_next=True), rpr(size=14, color=ACCENT, font=FONT_STRONG, caps=True, spacing=20)))
    s.append(pstyle('OpencdmpReferenceDetails', 'OpenCDMP Reference Details', ppr(before=0, after=120, line=276), rpr(size=16, color=MUTED)))
    s.append(pstyle('OpencdmpPlaceholder', 'OpenCDMP Placeholder', ppr(after=100), rpr(size=17, color=FAINT, italic=True)))
    s.append(pstyle('OpencdmpComment', 'OpenCDMP Comment', ppr(before=80, after=160, border=border('left', LINE, 18, 8), indent_left=180), rpr(size=17, color=MUTED)))
    s.append(pstyle('OpencdmpTableHeader', 'OpenCDMP Table Header', ppr(before=30, after=30, line=252), rpr(size=16, color=INK, font=FONT_STRONG)))
    s.append(pstyle('OpencdmpTableText', 'OpenCDMP Table Text', ppr(before=30, after=30, line=264), rpr(size=17, color=INK)))
    s.append(pstyle('OpencdmpFactLabel', 'OpenCDMP Fact Label', ppr(before=40, after=40), rpr(size=14, color=MUTED, font=FONT_STRONG, caps=True, spacing=16)))
    s.append(pstyle('OpencdmpFactValue', 'OpenCDMP Fact Value', ppr(before=20, after=20, line=276), rpr(size=19, color=INK)))
    s.append(pstyle('Caption', 'caption', ppr(before=60, after=240, align='center'), rpr(size=15, color=MUTED)))
    s.append(pstyle('TOC1', 'toc 1', ppr(before=140, after=40, tabs=TOC_TAB), rpr(size=19, color=INK, font=FONT_STRONG), next_style='Normal'))
    s.append(pstyle('TOC2', 'toc 2', ppr(before=20, after=20, tabs=TOC_TAB, indent_left=300), rpr(size=18, color=MUTED), next_style='Normal'))
    s.append(pstyle('Header', 'header', ppr(after=0), rpr(size=14, color=FAINT)))
    s.append(pstyle('Footer', 'footer', ppr(after=0, tabs='<w:tab w:val="right" w:pos="%d"/>' % CONTENT_W), rpr(size=14, color=FAINT)))
    s.append(cstyle('Hyperlink', 'Hyperlink', rpr(color=ACCENT, underline=True)))
    s.append(cstyle('OpencdmpStrongChar', 'OpenCDMP Strong Char', rpr(font=FONT_STRONG)))
    s.append(cstyle('OpencdmpNumberChar', 'OpenCDMP Number Char', rpr(color=ACCENT)))
    s.append(cstyle('OpencdmpMetaChar', 'OpenCDMP Meta Char', rpr(color=MUTED)))
    s.append(cstyle('OpencdmpHeaderKickerChar', 'OpenCDMP Header Kicker Char', rpr(size=13, color=ACCENT, font=FONT_STRONG, caps=True, spacing=24)))
    s.append(cstyle('OpencdmpPageNumberChar', 'OpenCDMP Page Number Char', rpr(color=ACCENT, font=FONT_STRONG)))
    none = ''.join(border(side, 'auto', 0) .replace('single', 'nil') for side in ('left', 'right', 'insideV'))
    s.append(tstyle('OpencdmpTableLined', 'OpenCDMP Table Lined',
                    border('top', LINE) + border('bottom', LINE) + border('insideH', LINE) + none, ACCENT_SOFT, header_bottom=ACCENT))
    s.append(tstyle('OpencdmpTableGrid', 'OpenCDMP Table Grid',
                    ''.join(border(side, LINE) for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')), ACCENT_SOFT))
    s.append(tstyle('OpencdmpTableStriped', 'OpenCDMP Table Striped',
                    border('bottom', LINE) + ''.join(border(side, 'auto', 0).replace('single', 'nil') for side in ('top', 'left', 'right', 'insideH', 'insideV')),
                    ACCENT_SOFT, band=STRIPE))
    s.append(tstyle('OpencdmpTableFacts', 'OpenCDMP Table Facts',
                    border('insideH', LINE) + border('bottom', LINE) + ''.join(border(side, 'auto', 0).replace('single', 'nil') for side in ('top', 'left', 'right', 'insideV')),
                    None, cell_margins=(90, 120, 90, 0), header_bold=False))
    defaults = ('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:eastAsia="{f}" w:cs="{f}"/>'
                '<w:sz w:val="19"/><w:szCs w:val="19"/><w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
                '<w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="324" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>').format(f=FONT)
    base = ('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
            '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/>'
            '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
            '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/></w:style>')
    normal = s.pop(0).replace('<w:style w:type="paragraph"', '<w:style w:type="paragraph" w:default="1"')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles %s>%s%s%s%s</w:styles>' % (W, defaults, normal, base, ''.join(s))


# ---------------------------------------------------------------- document parts


def run(text, char_style=None):
    style = '<w:rPr><w:rStyle w:val="%s"/></w:rPr>' % char_style if char_style else ''
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (style, escape(text))


def para(style, *runs, extra_ppr=''):
    style_xml = '<w:pStyle w:val="%s"/>' % style if style else ''
    return '<w:p><w:pPr>%s%s</w:pPr>%s</w:p>' % (style_xml, extra_ppr, ''.join(runs))


def tag(text, style='Normal'):
    """A paragraph holding only a block tag such as {{?chapters}}; poi-tl removes it when rendering."""
    return para(style, run(text))


def field(instruction, placeholder='', char_style=None):
    style = '<w:rPr><w:rStyle w:val="%s"/></w:rPr>' % char_style if char_style else ''
    return ('<w:r>{s}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{s}<w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
            '<w:r>{s}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{s}<w:t>%s</w:t></w:r><w:r>{s}<w:fldChar w:fldCharType="end"/></w:r>'
            % (instruction, placeholder)).format(s=style)


def header_xml(left, right):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr %s>%s</w:hdr>'
            % (W, para('Header', run(left), '<w:r><w:tab/></w:r>', run(right, 'OpencdmpHeaderKickerChar'),
                       extra_ppr='<w:tabs><w:tab w:val="right" w:pos="%d"/></w:tabs>' % CONTENT_W)))


def footer_xml(left):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr %s>%s</w:ftr>'
            % (W, para('Footer', run(left), '<w:r><w:tab/></w:r>', run('{{labels.page}} '), field('PAGE', '1', 'OpencdmpPageNumberChar'),
                       run(' {{labels.of}} '), field('NUMPAGES', '1'))))


def empty_hdr_ftr(kind):
    root = 'hdr' if kind == 'header' else 'ftr'
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:%s %s>%s</w:%s>' % (root, W, para(None), root)


def sect_pr():
    return ('<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:headerReference w:type="first" r:id="rIdHeaderFirst"/>'
            '<w:footerReference w:type="default" r:id="rIdFooter"/><w:footerReference w:type="first" r:id="rIdFooterFirst"/>'
            '<w:pgSz w:w="%d" w:h="%d"/><w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="624" w:footer="567" w:gutter="0"/>'
            '<w:cols w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>'
            % (PAGE_W, PAGE_H, MARGIN_TOP, MARGIN_X, MARGIN_BOTTOM, MARGIN_X))


COVER_STRIPE = ('<w:r><w:pict><v:rect id="OpencdmpCoverStripe" o:spid="_x0000_s3001" stroked="f" fillcolor="#%s" '
                'style="position:absolute;margin-left:0;margin-top:0;width:20pt;height:842pt;z-index:-251640832;'
                'mso-position-horizontal:left;mso-position-horizontal-relative:page;'
                'mso-position-vertical:top;mso-position-vertical-relative:page"/></w:pict></w:r>' % ACCENT)


def cover(kicker, title, subtitle, description, facts, entity):
    return ''.join([
        tag('{{?layout.cover}}'),
        para('Normal', run('{{@branding.logo}}'), extra_ppr='<w:spacing w:before="0" w:after="2200"/>'),
        # anchored to the kicker, which is always there (the logo paragraph is dropped when there is no logo)
        para('OpencdmpKicker', COVER_STRIPE, run(kicker)),
        tag('{{?%s.shortTitle}}' % entity),
        para('Title', run(title)),
        tag('{{/%s.shortTitle}}' % entity),
        tag('{{?%s.longTitle}}' % entity),
        para('OpencdmpTitleLong', run(title)),
        tag('{{/%s.longTitle}}' % entity),
        para('Subtitle', run(subtitle)),
        para('OpencdmpRule'),
        para('OpencdmpLead', run(description)),
        para('Normal', run(facts), extra_ppr='<w:spacing w:before="240"/>'),
        para('OpencdmpCoverNote', run('{{export.note}}')),
        para('Normal', run('{{>pageBreak}}')),
        tag('{{/layout.cover}}'),
    ])


def contents():
    return ''.join([
        tag('{{?layout.toc}}'),
        para('OpencdmpContentsTitle', run('{{labels.contents}}')),
        para('Normal', run('{{>toc}}')),
        para('Normal', run('{{>pageBreak}}')),
        tag('{{/layout.toc}}'),
    ])


def plan_body():
    return ''.join([
        cover('{{plan.kind}}', '{{plan.title}}',
              '{{labels.version}} {{plan.version}}  ·  {{plan.status}}  ·  {{labels.updated}} {{plan.updatedAt}}',
              '{{>plan.description}}', '{{>plan.facts}}', 'plan'),
        contents(),
        tag('{{?chapters}}'),
        tag('{{?group}}'),
        tag('{{?newPage}}'),
        para('Normal', run('{{>pageBreak}}')),
        tag('{{/newPage}}'),
        para('Heading1', run('{{numberPrefix}}', 'OpencdmpNumberChar'), run('{{title}}')),
        tag('{{/group}}'),
        tag('{{?section}}'),
        tag('{{?newPage}}'),
        para('Normal', run('{{>pageBreak}}')),
        tag('{{/newPage}}'),
        para('Heading1', run('{{numberPrefix}}', 'OpencdmpNumberChar'), run('{{title}}')),
        para('OpencdmpGuidance', run('{{>description}}')),
        para('Normal', run('{{>body}}')),
        tag('{{/section}}'),
        tag('{{?asset}}'),
        tag('{{?newPage}}'),
        para('Normal', run('{{>pageBreak}}')),
        tag('{{/newPage}}'),
        para('OpencdmpAssetTitle', run('{{numberPrefix}}', 'OpencdmpNumberChar'), run('{{title}}')),
        para('OpencdmpAssetMeta', run('{{meta}}')),
        para('OpencdmpAssetLead', run('{{>description}}')),
        para('Normal', run('{{>body}}')),
        tag('{{/asset}}'),
        tag('{{/chapters}}'),
    ])


def asset_body():
    return ''.join([
        cover('{{labels.asset}}', '{{asset.title}}', '{{labels.partOf}} {{plan.title}}',
              '{{>asset.description}}', '{{>asset.facts}}', 'asset'),
        # no table of contents by default: an asset is short; custom templates can add {{>toc}}
        para('Normal', run('{{>asset.body}}')),
    ])


def document_xml(body):
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s %s><w:body>%s%s</w:body></w:document>' % (W, DRAWING_NS, body, sect_pr())


# ---------------------------------------------------------------- fonts


FONTS = [  # (family, slot, file)
    (FONT, 'Regular', 'NotoSans-Regular.ttf'),
    (FONT, 'Bold', 'NotoSans-Bold.ttf'),
    (FONT, 'Italic', 'NotoSans-Italic.ttf'),
    (FONT, 'BoldItalic', 'NotoSans-BoldItalic.ttf'),
    (FONT_STRONG, 'Regular', 'NotoSans-SemiBold.ttf'),
]


def obfuscate(data, key):
    """ECMA-376 font obfuscation: the first 32 bytes XORed with the GUID key bytes in reverse order."""
    hex_key = key.replace('{', '').replace('}', '').replace('-', '')
    key_bytes = [int(hex_key[i:i + 2], 16) for i in range(0, 32, 2)][::-1]
    head = bytes(b ^ key_bytes[i % 16] for i, b in enumerate(data[:32]))
    return head + data[32:]


def font_parts(seed):
    """fontTable.xml, its relationships, and the obfuscated font files (deterministic keys per template)."""
    families = {}
    rels, files = [], {}
    for index, (family, slot, filename) in enumerate(FONTS, start=1):
        key = '{%s}' % str(uuid.uuid5(uuid.NAMESPACE_URL, 'opencdmp-%s-%s' % (seed, filename))).upper()
        with open(os.path.join(HERE, 'fonts', filename), 'rb') as f:
            files['word/fonts/font%d.odttf' % index] = obfuscate(f.read(), key)
        rels.append('<Relationship Id="rIdFont%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font%d.odttf"/>' % (index, index))
        families.setdefault(family, []).append('<w:embed%s r:id="rIdFont%d" w:fontKey="%s"/>' % (slot, index, key))
    table = ''.join('<w:font w:name="%s"><w:panose1 w:val="020B0502040504020204"/><w:charset w:val="00"/><w:family w:val="swiss"/>'
                    '<w:pitch w:val="variable"/>%s</w:font>' % (family, ''.join(embeds)) for family, embeds in families.items())
    font_table = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts %s>%s</w:fonts>' % (W, table)
    font_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>' % ''.join(rels)
    return font_table, font_rels, files


# ---------------------------------------------------------------- package


def settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings %s><w:embedTrueTypeFonts/><w:saveSubsetFonts/>'
            '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' % W)


def build(name, title, body, header_left, header_right, footer_left):
    font_table, font_rels, font_files = font_parts(name)
    parts = {
        '[Content_Types].xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
            '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
            + ''.join('<Override PartName="/word/%s.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml"/>'
                      % (p, 'header' if p.startswith('header') else 'footer') for p in ('header1', 'header2', 'footer1', 'footer2')) +
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '</Types>'),
        '_rels/.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>'),
        'word/_rels/document.xml.rels': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
            '<Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
            '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
            '<Relationship Id="rIdHeaderFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>'
            '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
            '<Relationship Id="rIdFooterFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>'
            '</Relationships>'),
        'word/_rels/fontTable.xml.rels': font_rels,
        'word/document.xml': document_xml(body),
        'word/styles.xml': styles_xml(),
        'word/settings.xml': settings_xml(),
        'word/fontTable.xml': font_table,
        'word/header1.xml': header_xml(header_left, header_right),
        'word/header2.xml': empty_hdr_ftr('header'),
        'word/footer1.xml': footer_xml(footer_left),
        'word/footer2.xml': empty_hdr_ftr('footer'),
        'docProps/core.xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>OpenCDMP</dc:creator></cp:coreProperties>' % title),
        'docProps/app.xml': (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
            '<Application>OpenCDMP</Application></Properties>'),
    }
    path = os.path.join(OUT, name + '.docx')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        for part, content in parts.items():
            info = zipfile.ZipInfo(part, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content.encode('utf-8') if isinstance(content, str) else content)
        for part, content in font_files.items():
            info = zipfile.ZipInfo(part, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, content)
    print('wrote', path, os.path.getsize(path) // 1024, 'KB')


if __name__ == '__main__':
    build('plan', 'OpenCDMP plan template', plan_body(), '{{plan.title}}', '{{plan.kind}}', '{{plan.status}}  ·  {{labels.version}} {{plan.version}}')
    build('description', 'OpenCDMP asset template', asset_body(), '{{asset.title}}', '{{plan.title}}', '{{asset.status}}')
