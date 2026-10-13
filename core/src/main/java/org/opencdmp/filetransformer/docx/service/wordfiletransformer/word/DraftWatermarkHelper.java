package org.opencdmp.filetransformer.docx.service.wordfiletransformer.word;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * Stamps a diagonal text watermark (e.g. DRAFT) on every page by appending a VML shape to each
 * header. Existing template headers are preserved — the watermark paragraph is appended to them
 * rather than replacing them, so custom uploaded templates keep their header content.
 */
public final class DraftWatermarkHelper {
    private static final Logger logger = LoggerFactory.getLogger(DraftWatermarkHelper.class);

    private DraftWatermarkHelper() {
    }

    /**
     * @param color   VML fill colour, e.g. {@code "#c0c0c0"}
     * @param opacity VML fill opacity, e.g. {@code ".3"}; LibreOffice (the PDF conversion) ignores it,
     *                so a light colour at full opacity looks the same in Word and in the PDF
     * @param width   width of the diagonal text box in points; the height scales the letters
     */
    public static void applyWatermark(XWPFDocument document, String text, String fontFamily, String color, String opacity, double width, double height) {
        try {
            List<XWPFHeader> headers = new ArrayList<>(document.getHeaderList() != null ? document.getHeaderList() : List.of());
            if (headers.isEmpty()) {
                // template without headers: create a default one and use the returned instance
                // directly — a freshly created header is not guaranteed to show up in getHeaderList()
                XWPFHeaderFooterPolicy policy = document.getHeaderFooterPolicy();
                if (policy == null) policy = document.createHeaderFooterPolicy();
                headers.add(policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT));
            }
            int shapeId = 2025;
            for (XWPFHeader header : headers) {
                appendWatermarkParagraph(header, text, shapeId++, fontFamily, color, opacity, width, height);
            }
        } catch (Exception e) {
            // the watermark is decoration — never fail an export because of it
            logger.error("could not apply watermark", e);
        }
    }

    private static void appendWatermarkParagraph(XWPFHeader header, String text, int shapeId, String fontFamily, String color, String opacity, double width, double height) throws Exception {
        String escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;");
        String xml =
                "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                        "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                        "xmlns:o=\"urn:schemas-microsoft-com:office:office\">" +
                        "<w:pPr><w:pStyle w:val=\"Header\"/></w:pPr>" +
                        "<w:r><w:rPr><w:noProof/></w:rPr><w:pict>" +
                        "<v:shapetype id=\"_x0000_t136\" coordsize=\"21600,21600\" o:spt=\"136\" adj=\"10800\" path=\"m@7,l@8,m@5,21600l@6,21600e\">" +
                        "<v:formulas>" +
                        "<v:f eqn=\"sum #0 0 10800\"/>" +
                        "<v:f eqn=\"prod #0 2 1\"/>" +
                        "<v:f eqn=\"sum 21600 0 @1\"/>" +
                        "<v:f eqn=\"sum 0 0 @2\"/>" +
                        "<v:f eqn=\"sum 21600 0 @3\"/>" +
                        "<v:f eqn=\"if @0 @3 0\"/>" +
                        "<v:f eqn=\"if @0 21600 @1\"/>" +
                        "<v:f eqn=\"if @0 0 @2\"/>" +
                        "<v:f eqn=\"if @0 @4 21600\"/>" +
                        "<v:f eqn=\"mid @5 @6\"/>" +
                        "<v:f eqn=\"mid @8 @5\"/>" +
                        "<v:f eqn=\"mid @7 @8\"/>" +
                        "<v:f eqn=\"mid @6 @7\"/>" +
                        "<v:f eqn=\"sum @6 0 @5\"/>" +
                        "</v:formulas>" +
                        "<v:path textpathok=\"t\" o:connecttype=\"custom\" o:connectlocs=\"@9,0;@10,10800;@11,21600;@12,10800\" o:connectangles=\"270,180,90,0\"/>" +
                        "<v:textpath on=\"t\" fitshape=\"t\"/>" +
                        "</v:shapetype>" +
                        "<v:shape id=\"OpencdmpWatermark" + shapeId + "\" o:spid=\"_x0000_s" + shapeId + "\" " +
                        "type=\"#_x0000_t136\" " +
                        "style=\"position:absolute;margin-left:0;margin-top:0;width:" + String.format(java.util.Locale.ROOT, "%.2f", width) + "pt;height:" + String.format(java.util.Locale.ROOT, "%.2f", height) + "pt;rotation:315;" +
                        "z-index:-251654144;mso-wrap-edited:f;" +
                        "mso-position-horizontal:center;mso-position-horizontal-relative:margin;" +
                        "mso-position-vertical:center;mso-position-vertical-relative:margin\" " +
                        "o:allowincell=\"f\" fillcolor=\"" + color + "\" stroked=\"f\">" +
                        "<v:fill opacity=\"" + opacity + "\"/>" +
                        "<v:textpath style=\"font-family:&quot;" + fontFamily + "&quot;;font-size:1pt\" string=\"" + escaped + "\"/>" +
                        "</v:shape>" +
                        "</w:pict></w:r></w:p>";
        CTP watermarkCtp = CTP.Factory.parse(xml);
        XWPFParagraph paragraph = header.createParagraph();
        try (XmlCursor target = paragraph.getCTP().newCursor(); XmlCursor source = watermarkCtp.newCursor()) {
            source.moveXmlContents(target);
        }
    }
}
