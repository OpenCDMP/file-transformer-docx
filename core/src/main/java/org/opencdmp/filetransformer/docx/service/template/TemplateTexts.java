package org.opencdmp.filetransformer.docx.service.template;

import org.apache.poi.xwpf.usermodel.*;

import java.util.ArrayList;
import java.util.List;

/** Reads the text of a DOCX template, to check it before it is rendered. */
public final class TemplateTexts {

    /** The start of the placeholders of the original engine, which is no longer supported. */
    public static final String OLD_PLACEHOLDER = "'{OPENCDMP.";

    private TemplateTexts() {
    }

    /** Whether the template was made for the original engine: it uses {@code '{OPENCDMP.X}'} placeholders. */
    public static boolean hasOldPlaceholders(XWPFDocument document) {
        return texts(document).stream().anyMatch(x -> x.contains(OLD_PLACEHOLDER));
    }

    /** The text of every paragraph in the body, tables (including nested), headers and footers. */
    public static List<String> texts(XWPFDocument document) {
        List<String> texts = new ArrayList<>();
        collect(document.getBodyElements(), texts);
        for (XWPFHeader header : document.getHeaderList()) collect(header.getBodyElements(), texts);
        for (XWPFFooter footer : document.getFooterList()) collect(footer.getBodyElements(), texts);
        return texts;
    }

    private static void collect(List<IBodyElement> elements, List<String> texts) {
        for (IBodyElement element : elements) {
            if (element instanceof XWPFParagraph paragraph) {
                texts.add(paragraph.getText());
            } else if (element instanceof XWPFTable table) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) collect(cell.getBodyElements(), texts);
                }
            }
        }
    }
}
