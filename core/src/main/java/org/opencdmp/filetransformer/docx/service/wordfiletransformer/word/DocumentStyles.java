package org.opencdmp.filetransformer.docx.service.wordfiletransformer.word;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;

/** Document-wide adjustments applied to a rendered export: the accent colour and the page size. */
public final class DocumentStyles {
    private static final Logger logger = LoggerFactory.getLogger(DocumentStyles.class);

    private DocumentStyles() {
    }

    /**
     * Rebrands a rendered document: every use of the template's accent colour, and of its
     * table-header tint, is rewritten to the configured accent (and its tint).
     */
    public static void recolor(XWPFDocument document, String templateAccent, String accentColor) {
        if (accentColor == null || templateAccent == null || accentColor.equalsIgnoreCase(templateAccent)) return;
        java.util.Map<String, String> colors = java.util.Map.of(
                templateAccent.toUpperCase(), accentColor.toUpperCase(),
                tint(templateAccent, 0.10), tint(accentColor, 0.10));
        try {
            if (document.getStyles() != null && document.getStyles().getCtStyles() != null) replaceColors(document.getStyles().getCtStyles(), colors);
            replaceColors(document.getDocument(), colors);
            document.getHeaderList().forEach(h -> replaceColors(h._getHdrFtr(), colors));
            document.getFooterList().forEach(f -> replaceColors(f._getHdrFtr(), colors));
        } catch (Exception e) {
            logger.error("could not apply accent color, keeping the template colors", e);
        }
    }

    private static final javax.xml.namespace.QName VML_FILL = new javax.xml.namespace.QName("", "fillcolor");

    private static void replaceColors(org.apache.xmlbeans.XmlObject part, java.util.Map<String, String> colors) {
        String ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        javax.xml.namespace.QName[] attributes = {new javax.xml.namespace.QName(ns, "val"), new javax.xml.namespace.QName(ns, "color"), new javax.xml.namespace.QName(ns, "fill")};
        try (org.apache.xmlbeans.XmlCursor cursor = part.newCursor()) {
            while (cursor.hasNextToken()) {
                if (!cursor.toNextToken().isStart()) continue;
                for (javax.xml.namespace.QName attribute : attributes) {
                    String value = cursor.getAttributeText(attribute);
                    if (value == null) continue;
                    String replacement = colors.get(value.toUpperCase());
                    if (replacement != null) cursor.setAttributeText(attribute, replacement);
                }
                // VML shapes: fillcolor="#2457C5"
                String fill = cursor.getAttributeText(VML_FILL);
                if (fill != null && fill.startsWith("#")) {
                    String replacement = colors.get(fill.substring(1).toUpperCase());
                    if (replacement != null) cursor.setAttributeText(VML_FILL, "#" + replacement);
                }
            }
        }
    }

    /** Sets the page size on the document's main section (portrait A4 or Letter). */
    public static void applyPageSize(XWPFDocument document, boolean a4) {
        try {
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr sectPr =
                    document.getDocument().getBody().isSetSectPr()
                            ? document.getDocument().getBody().getSectPr()
                            : document.getDocument().getBody().addNewSectPr();
            org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz pgSz =
                    sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
            pgSz.setW(BigInteger.valueOf(a4 ? 11906 : 12240));
            pgSz.setH(BigInteger.valueOf(a4 ? 16838 : 15840));
        } catch (Exception e) {
            logger.error("could not apply page size, keeping template default", e);
        }
    }

    /** Mixes a hex color towards white; factor 0.12 keeps 12% of the color. */
    public static String tint(String hex, double factor) {
        try {
            int r = Integer.parseInt(hex.substring(0, 2), 16);
            int g = Integer.parseInt(hex.substring(2, 4), 16);
            int b = Integer.parseInt(hex.substring(4, 6), 16);
            r = (int) Math.round(255 - (255 - r) * factor);
            g = (int) Math.round(255 - (255 - g) * factor);
            b = (int) Math.round(255 - (255 - b) * factor);
            return String.format("%02X%02X%02X", r, g, b);
        } catch (Exception e) {
            return hex.toUpperCase();
        }
    }
}
