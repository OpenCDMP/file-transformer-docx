package org.opencdmp.filetransformer.docx.service.template.render;

import com.deepoove.poi.xwpf.BodyContainer;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

/** Where generated content goes: before a template placeholder, or into a table cell. */
public interface Sink {

    XWPFParagraph paragraph();

    /** A new table, or {@code null} where tables are not supported (inside a table cell). */
    XWPFTable table(int rows, int columns);

    boolean inTable();

    /** Inserts before the paragraph holding a placeholder run, in call order. */
    final class BeforePlaceholder implements Sink {
        private final BodyContainer container;
        private final XWPFRun placeholder;

        public BeforePlaceholder(BodyContainer container, XWPFRun placeholder) {
            this.container = container;
            this.placeholder = placeholder;
        }

        @Override
        public XWPFParagraph paragraph() {
            return this.container.insertNewParagraph(this.placeholder);
        }

        @Override
        public XWPFTable table(int rows, int columns) {
            return this.container.insertNewTable(this.placeholder, rows, columns);
        }

        @Override
        public boolean inTable() {
            return false;
        }
    }

    /** Appends to a table cell, reusing the empty paragraph every new cell starts with. */
    final class InCell implements Sink {
        private final XWPFTableCell cell;
        private boolean used;

        public InCell(XWPFTableCell cell) {
            this.cell = cell;
        }

        @Override
        public XWPFParagraph paragraph() {
            if (!this.used) {
                this.used = true;
                if (!this.cell.getParagraphs().isEmpty() && this.cell.getParagraphs().getFirst().getRuns().isEmpty()) return this.cell.getParagraphs().getFirst();
            }
            return this.cell.addParagraph();
        }

        @Override
        public XWPFTable table(int rows, int columns) {
            return null;
        }

        @Override
        public boolean inTable() {
            return true;
        }
    }
}
