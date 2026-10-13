package org.opencdmp.filetransformer.docx.service.template.render;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.FactView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordExportOptions;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** Writes {@link Body} content (sections, field sets, questions, answers) and reference lists. */
final class BodyWriter {

    private final DocxWriter writer;
    private final HtmlRenderer html;
    private final Map<String, String> labels;
    private final WordExportOptions.ReferenceStyle referenceStyle;
    private final String tableStyle;

    BodyWriter(DocxWriter writer, Map<String, String> labels, WordExportOptions.ReferenceStyle referenceStyle, WordExportOptions.TableStyle tableStyle) {
        this.writer = writer;
        this.html = new HtmlRenderer(writer);
        this.labels = labels;
        this.referenceStyle = referenceStyle;
        this.tableStyle = switch (tableStyle) {
            case LINED -> TemplateStyles.TABLE_LINED;
            case GRID -> TemplateStyles.TABLE_GRID;
            case STRIPED -> TemplateStyles.TABLE_STRIPED;
        };
    }

    void body(Body body, Sink sink) {
        this.nodes(body.nodes(), body.baseLevel(), sink);
    }

    private void nodes(List<Body.Node> nodes, int level, Sink sink) {
        for (Body.Node node : nodes) {
            switch (node) {
                case Body.SectionNode section -> {
                    XWPFParagraph heading = DocxWriter.styled(sink.paragraph(), TemplateStyles.heading(level));
                    this.numbered(heading, section.number(), section.title());
                    if (section.description() != null) this.html.render(section.description().html(), sink, TemplateStyles.GUIDANCE);
                    this.nodes(section.children(), level + 1, sink);
                }
                case Body.FieldSetNode fieldSet -> this.fieldSet(fieldSet, sink);
            }
        }
    }

    private void numbered(XWPFParagraph paragraph, String number, String title) {
        if (number != null) {
            // number and spacing share a run: Word's TOC drops whitespace-only runs
            DocxWriter.text(paragraph, number + " ", TemplateStyles.NUMBER_CHAR);
        }
        DocxWriter.text(paragraph, title != null ? title : "", null);
    }

    private void fieldSet(Body.FieldSetNode fieldSet, Sink sink) {
        if (fieldSet.title() != null && !fieldSet.title().isBlank()) {
            XWPFParagraph title = DocxWriter.styled(sink.paragraph(), TemplateStyles.FIELD_SET_TITLE);
            this.numbered(title, fieldSet.number(), fieldSet.title());
        }
        if (fieldSet.description() != null) this.html.render(fieldSet.description().html(), sink, TemplateStyles.GUIDANCE);
        switch (fieldSet.mode()) {
            case TABLE -> this.itemsTable(fieldSet, sink);
            case TRANSPOSED -> this.transposedTable(fieldSet, sink);
            default -> {
                boolean several = fieldSet.items().size() > 1;
                int index = 0;
                for (Body.Item item : fieldSet.items()) {
                    index++;
                    if (several) DocxWriter.text(DocxWriter.styled(sink.paragraph(), TemplateStyles.ENTRY), this.labels.get("entry") + " " + index, null);
                    if (fieldSet.mode() == Body.Mode.QUESTION_TABLE && !sink.inTable()) this.questionTable(item, sink);
                    else this.questions(item, sink, fieldSet.mode() == Body.Mode.COMPACT);
                }
            }
        }
        if (fieldSet.comment() != null) {
            XWPFParagraph comment = DocxWriter.styled(sink.paragraph(), TemplateStyles.COMMENT);
            DocxWriter.text(comment, this.labels.get("comment") + ": ", TemplateStyles.STRONG_CHAR);
            DocxWriter.text(comment, fieldSet.comment(), null);
        }
    }

    private void questions(Body.Item item, Sink sink, boolean compact) {
        for (Body.Question question : item.questions()) {
            if (!compact && question.label() != null) {
                XWPFParagraph label = DocxWriter.styled(sink.paragraph(), TemplateStyles.QUESTION);
                DocxWriter.keepWithNext(label);
                DocxWriter.text(label, question.label(), null);
            }
            this.answer(question.answer(), sink, TemplateStyles.ANSWER);
        }
    }

    private void questionTable(Body.Item item, Sink sink) {
        XWPFTable table = sink.table(item.questions().size(), 2);
        if (table == null) {
            this.questions(item, sink, false);
            return;
        }
        for (int i = 0; i < item.questions().size(); i++) {
            Body.Question question = item.questions().get(i);
            DocxWriter.text(DocxWriter.styled(new Sink.InCell(table.getRow(i).getCell(0)).paragraph(), TemplateStyles.QUESTION), question.label() != null ? question.label() : "", null);
            this.answer(question.answer(), new Sink.InCell(table.getRow(i).getCell(1)), TemplateStyles.TABLE_TEXT);
        }
        this.writer.table(table, this.tableStyle, new double[]{0.36, 0.64}, false);
        this.spacer(sink);
    }

    private void itemsTable(Body.FieldSetNode fieldSet, Sink sink) {
        int columns = fieldSet.columns().size();
        XWPFTable table = columns == 0 ? null : sink.table(fieldSet.items().size() + 1, columns);
        if (table == null) {
            fieldSet.items().forEach(item -> this.questions(item, sink, false));
            return;
        }
        for (int c = 0; c < columns; c++) {
            DocxWriter.text(DocxWriter.styled(new Sink.InCell(table.getRow(0).getCell(c)).paragraph(), TemplateStyles.TABLE_HEADER), fieldSet.columns().get(c), null);
        }
        for (int r = 0; r < fieldSet.items().size(); r++) {
            List<Body.Question> questions = fieldSet.items().get(r).questions();
            for (int c = 0; c < columns; c++) {
                Sink cell = new Sink.InCell(table.getRow(r + 1).getCell(c));
                if (c < questions.size()) this.answer(questions.get(c).answer(), cell, TemplateStyles.TABLE_TEXT);
                else DocxWriter.styled(cell.paragraph(), TemplateStyles.TABLE_TEXT);
            }
        }
        double[] widths = new double[columns];
        Arrays.fill(widths, 1);
        this.writer.table(table, this.tableStyle, widths, true);
        this.spacer(sink);
    }

    private void transposedTable(Body.FieldSetNode fieldSet, Sink sink) {
        int rows = fieldSet.columns().size();
        int items = fieldSet.items().size();
        XWPFTable table = rows == 0 ? null : sink.table(rows, items + 1);
        if (table == null) {
            fieldSet.items().forEach(item -> this.questions(item, sink, false));
            return;
        }
        for (int r = 0; r < rows; r++) {
            DocxWriter.text(DocxWriter.styled(new Sink.InCell(table.getRow(r).getCell(0)).paragraph(), TemplateStyles.QUESTION), fieldSet.columns().get(r), null);
            for (int i = 0; i < items; i++) {
                List<Body.Question> questions = fieldSet.items().get(i).questions();
                Sink cell = new Sink.InCell(table.getRow(r).getCell(i + 1));
                if (r < questions.size()) this.answer(questions.get(r).answer(), cell, TemplateStyles.TABLE_TEXT);
                else DocxWriter.styled(cell.paragraph(), TemplateStyles.TABLE_TEXT);
            }
        }
        double[] widths = new double[items + 1];
        Arrays.fill(widths, 1);
        widths[0] = Math.max(1, items * 0.5);
        this.writer.table(table, this.tableStyle, widths, false);
        this.spacer(sink);
    }

    /** Space after a table, which otherwise touches the next heading. */
    private void spacer(Sink sink) {
        XWPFParagraph spacer = sink.paragraph();
        spacer.setSpacingAfter(0);
        spacer.setSpacingBefore(0);
    }

    void answer(Body.Answer answer, Sink sink, String style) {
        switch (answer) {
            case Body.Text text -> DocxWriter.text(DocxWriter.styled(sink.paragraph(), style), text.value(), null);
            case Body.Html html -> this.html.render(html.html(), sink, style);
            case Body.Values values -> {
                for (String value : values.values()) {
                    XWPFParagraph item = DocxWriter.styled(sink.paragraph(), style);
                    this.writer.bullet(item, 0);
                    item.setSpacingAfter(40);
                    DocxWriter.text(item, value, null);
                }
            }
            case Body.References references -> this.references(references.references(), sink, style);
            case Body.Image image -> {
                if (!this.writer.picture(sink, image.data(), image.mimeType(), image.filename(), this.labels.get("image"))) {
                    DocxWriter.text(DocxWriter.styled(sink.paragraph(), style), image.filename(), null);
                }
            }
            case Body.Attachment attachment -> {
                XWPFParagraph paragraph = DocxWriter.styled(sink.paragraph(), style);
                DocxWriter.text(paragraph, this.labels.get("attachment") + ": ", TemplateStyles.META_CHAR);
                DocxWriter.text(paragraph, attachment.filename(), null);
            }
            case Body.Missing missing -> DocxWriter.text(DocxWriter.styled(sink.paragraph(), sink.inTable() ? style : TemplateStyles.PLACEHOLDER),
                    sink.inTable() ? "—" : this.labels.get("notAnswered"), sink.inTable() ? TemplateStyles.META_CHAR : null);
        }
    }

    void references(List<ReferenceView> references, Sink sink, String style) {
        if (references == null || references.isEmpty()) return;
        WordExportOptions.ReferenceStyle mode = this.referenceStyle;
        if (mode == WordExportOptions.ReferenceStyle.TABLE && (sink.inTable() || !this.referencesTable(references, sink))) mode = WordExportOptions.ReferenceStyle.DETAILED;
        if (mode == WordExportOptions.ReferenceStyle.TABLE) return;

        boolean several = references.size() > 1;
        for (ReferenceView reference : references) {
            XWPFParagraph title = DocxWriter.styled(sink.paragraph(), style);
            if (several) this.writer.bullet(title, 0);
            boolean details = mode == WordExportOptions.ReferenceStyle.DETAILED && !reference.details.isEmpty();
            title.setSpacingAfter(details ? 0 : (several ? 40 : 120));
            if (details) DocxWriter.keepWithNext(title);
            if (reference.url != null) DocxWriter.link(title, reference.url, reference.label);
            else DocxWriter.text(title, reference.label, TemplateStyles.STRONG_CHAR);
            if (!details) continue;

            XWPFParagraph line = DocxWriter.styled(sink.paragraph(), TemplateStyles.REFERENCE_DETAILS);
            if (several) line.setIndentationLeft(360);
            for (int i = 0; i < reference.details.size(); i++) {
                FactView detail = reference.details.get(i);
                if (i > 0) DocxWriter.text(line, "  ·  ", null);
                DocxWriter.text(line, detail.label + " ", TemplateStyles.META_CHAR);
                if (detail.url != null) DocxWriter.link(line, detail.url, detail.value);
                else DocxWriter.text(line, detail.value, null);
            }
        }
    }

    /** References as a table: the label plus one column per detail shown by any of them. */
    private boolean referencesTable(List<ReferenceView> references, Sink sink) {
        Set<String> columns = new LinkedHashSet<>();
        references.forEach(reference -> reference.details.forEach(detail -> columns.add(detail.label)));
        List<String> header = new ArrayList<>();
        header.add(references.getFirst().type != null ? references.getFirst().type : "");
        header.addAll(columns);
        XWPFTable table = sink.table(references.size() + 1, header.size());
        if (table == null) return false;
        for (int c = 0; c < header.size(); c++) {
            DocxWriter.text(DocxWriter.styled(new Sink.InCell(table.getRow(0).getCell(c)).paragraph(), TemplateStyles.TABLE_HEADER), header.get(c), null);
        }
        for (int r = 0; r < references.size(); r++) {
            ReferenceView reference = references.get(r);
            XWPFParagraph name = DocxWriter.styled(new Sink.InCell(table.getRow(r + 1).getCell(0)).paragraph(), TemplateStyles.TABLE_TEXT);
            if (reference.url != null) DocxWriter.link(name, reference.url, reference.label);
            else DocxWriter.text(name, reference.label, null);
            int c = 1;
            for (String column : columns) {
                XWPFParagraph cell = DocxWriter.styled(new Sink.InCell(table.getRow(r + 1).getCell(c++)).paragraph(), TemplateStyles.TABLE_TEXT);
                FactView detail = reference.details.stream().filter(x -> column.equals(x.label)).findFirst().orElse(null);
                if (detail == null) continue;
                if (detail.url != null) DocxWriter.link(cell, detail.url, detail.value);
                else DocxWriter.text(cell, detail.value, null);
            }
        }
        double[] widths = new double[header.size()];
        Arrays.fill(widths, 1);
        widths[0] = 1.6;
        this.writer.table(table, this.tableStyle, widths, true);
        this.spacer(sink);
        return true;
    }

    void facts(List<FactView> facts, Sink sink) {
        if (facts == null || facts.isEmpty()) return;
        XWPFTable table = sink.table(facts.size(), 2);
        if (table == null) {
            for (FactView fact : facts) {
                XWPFParagraph paragraph = DocxWriter.styled(sink.paragraph(), TemplateStyles.FACT_VALUE);
                DocxWriter.text(paragraph, fact.label + ": ", TemplateStyles.META_CHAR);
                DocxWriter.text(paragraph, fact.value, null);
            }
            return;
        }
        for (int r = 0; r < facts.size(); r++) {
            FactView fact = facts.get(r);
            DocxWriter.text(DocxWriter.styled(new Sink.InCell(table.getRow(r).getCell(0)).paragraph(), TemplateStyles.FACT_LABEL), fact.label, null);
            XWPFParagraph value = DocxWriter.styled(new Sink.InCell(table.getRow(r).getCell(1)).paragraph(), TemplateStyles.FACT_VALUE);
            if (fact.url != null) DocxWriter.link(value, fact.url, fact.value);
            else DocxWriter.text(value, fact.value, null);
        }
        this.writer.table(table, TemplateStyles.TABLE_FACTS, new double[]{0.28, 0.72}, false);
    }
}
