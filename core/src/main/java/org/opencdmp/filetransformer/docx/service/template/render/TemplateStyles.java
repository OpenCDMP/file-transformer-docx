package org.opencdmp.filetransformer.docx.service.template.render;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * The style contract of '{{...}}' templates. Generated content refers to these style ids only, so a
 * template restyles every generated paragraph, list and table by (re)defining them in Word.
 * A template that lacks one gets the definition of the built-in template, copied in at export time.
 */
public final class TemplateStyles {
    private static final Logger logger = LoggerFactory.getLogger(TemplateStyles.class);

    // paragraph styles
    public static final String NORMAL = "Normal";
    public static final String HEADING_PREFIX = "Heading";
    public static final String QUESTION = "OpencdmpQuestion";
    public static final String ANSWER = "OpencdmpAnswer";
    public static final String FIELD_SET_TITLE = "OpencdmpFieldSetTitle";
    public static final String ENTRY = "OpencdmpEntry";
    public static final String GUIDANCE = "OpencdmpGuidance";
    public static final String REFERENCE_DETAILS = "OpencdmpReferenceDetails";
    public static final String PLACEHOLDER = "OpencdmpPlaceholder";
    public static final String COMMENT = "OpencdmpComment";
    public static final String TABLE_HEADER = "OpencdmpTableHeader";
    public static final String TABLE_TEXT = "OpencdmpTableText";
    public static final String FACT_LABEL = "OpencdmpFactLabel";
    public static final String FACT_VALUE = "OpencdmpFactValue";
    public static final String CAPTION = "Caption";
    public static final String TOC_PREFIX = "TOC";

    // character styles
    public static final String HYPERLINK = "Hyperlink";
    public static final String STRONG_CHAR = "OpencdmpStrongChar";
    public static final String NUMBER_CHAR = "OpencdmpNumberChar";
    public static final String META_CHAR = "OpencdmpMetaChar";

    // table styles
    public static final String TABLE_LINED = "OpencdmpTableLined";
    public static final String TABLE_GRID = "OpencdmpTableGrid";
    public static final String TABLE_STRIPED = "OpencdmpTableStriped";
    public static final String TABLE_FACTS = "OpencdmpTableFacts";

    /** The accent colour of the built-in templates; the configured accent colour replaces it (and its 10% tint). */
    public static final String ACCENT_COLOR = "2457C5";
    /** Font of the draft watermark on '{{...}}' template exports. */
    public static final String WATERMARK_FONT = "Noto Sans SemiBold";
    /** Colour of the draft watermark, at full opacity (the PDF conversion ignores VML opacity). */
    public static final String WATERMARK_COLOR = "#F1F3F6";
    /** Size of the diagonal draft mark on template exports, in points: a quiet stamp rather than a page-wide banner. */
    public static final double WATERMARK_WIDTH = 380;
    public static final double WATERMARK_HEIGHT = 92;

    public static final List<String> CONTRACT = List.of(
            "Heading1", "Heading2", "Heading3", "Heading4", "Heading5", "Heading6", "Title", "Subtitle",
            QUESTION, ANSWER, FIELD_SET_TITLE, ENTRY, GUIDANCE, REFERENCE_DETAILS, PLACEHOLDER, COMMENT, TABLE_HEADER, TABLE_TEXT,
            FACT_LABEL, FACT_VALUE, CAPTION, "TOC1", "TOC2", HYPERLINK, STRONG_CHAR, NUMBER_CHAR, META_CHAR,
            TABLE_LINED, TABLE_GRID, TABLE_STRIPED, TABLE_FACTS);

    private TemplateStyles() {
    }

    /** Copies every contract style missing from {@code document} out of {@code reference} (the built-in template). */
    public static void ensure(XWPFDocument document, XWPFDocument reference) {
        if (reference == null) return;
        try {
            XWPFStyles target = document.createStyles();
            XWPFStyles source = reference.getStyles();
            if (source == null) return;
            for (String id : CONTRACT) {
                if (target.styleExist(id)) continue;
                XWPFStyle style = source.getStyle(id);
                if (style == null) continue;
                // copy the style's basedOn chain first so inherited formatting survives
                String basedOn = style.getBasisStyleID();
                if (basedOn != null && !target.styleExist(basedOn) && source.getStyle(basedOn) != null) {
                    target.addStyle(new XWPFStyle((CTStyle) source.getStyle(basedOn).getCTStyle().copy(), target));
                }
                target.addStyle(new XWPFStyle((CTStyle) style.getCTStyle().copy(), target));
            }
        } catch (Exception e) {
            // a malformed styles part in a custom template must not fail the export
            logger.error("could not ensure template styles, generated content falls back to the template's defaults", e);
        }
    }

    public static String heading(int level) {
        return HEADING_PREFIX + Math.max(1, Math.min(6, level));
    }
}
