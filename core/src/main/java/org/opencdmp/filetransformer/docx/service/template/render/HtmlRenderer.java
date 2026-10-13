package org.opencdmp.filetransformer.docx.service.template.render;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Set;

/**
 * Renders editor HTML into Word paragraphs styled with the given paragraph style: paragraphs,
 * line breaks, bold / italic / underline / strike / sub / sup / highlight / code, links, nested
 * bulleted and numbered lists, headings, block quotes and simple tables (flattened inside cells).
 * Colours, fonts and sizes from the HTML are ignored on purpose: the template decides the look.
 */
final class HtmlRenderer {

    private static final Set<String> BLOCKS = Set.of("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "ul", "ol", "li", "table", "hr", "section", "article", "header", "footer", "figure");

    private final DocxWriter writer;

    HtmlRenderer(DocxWriter writer) {
        this.writer = writer;
    }

    void render(String html, Sink sink, String paragraphStyle) {
        if (html == null || html.isBlank()) return;
        if (!html.contains("<")) {
            // plain text: blank lines separate paragraphs, single newlines break lines
            for (String block : html.trim().split("\\n\\s*\\n")) {
                DocxWriter.text(DocxWriter.styled(sink.paragraph(), paragraphStyle), block.trim(), null);
            }
            return;
        }
        Context context = new Context(sink, paragraphStyle);
        this.blocks(Jsoup.parseBodyFragment(html).body(), context, new Format());
        context.close();
    }

    private void blocks(Element parent, Context context, Format format) {
        for (Node node : parent.childNodes()) {
            if (node instanceof Element element && BLOCKS.contains(element.tagName())) {
                this.block(element, context, format);
            } else {
                this.inline(node, context, format);
            }
        }
    }

    private void block(Element element, Context context, Format format) {
        String tag = element.tagName();
        switch (tag) {
            case "ul", "ol" -> {
                context.close();
                BigInteger numId = tag.equals("ol") ? this.writer.newNumberedList() : null;
                context.listLevel++;
                for (Element item : element.children()) {
                    if (!item.tagName().equals("li")) continue;
                    context.listItem = numId == null ? ListItem.BULLET : ListItem.NUMBER;
                    context.listNumId = numId;
                    this.blocks(item, context, format);
                    context.close();
                }
                context.listLevel--;
                context.listItem = null;
            }
            case "li" -> {
                this.blocks(element, context, format);
                context.close();
            }
            case "table" -> this.table(element, context, format);
            case "hr" -> context.close();
            case "h1", "h2", "h3", "h4", "h5", "h6" -> {
                context.close();
                XWPFParagraph paragraph = context.open();
                DocxWriter.keepWithNext(paragraph);
                Format heading = format.copy();
                heading.bold = true;
                this.blocks(element, context, heading);
                context.close();
            }
            case "blockquote" -> {
                context.close();
                context.quote++;
                this.blocks(element, context, format);
                context.close();
                context.quote--;
            }
            case "pre" -> {
                context.close();
                Format code = format.copy();
                code.code = true;
                context.open();
                context.run(element.wholeText(), code);
                context.close();
            }
            default -> {
                // p, div, ...: a paragraph of its own; list items reuse the item's first paragraph
                if (context.listItem == null || context.paragraph != null) context.close();
                XWPFParagraph paragraph = context.open();
                String align = alignment(element);
                if (align != null) paragraph.setAlignment(switch (align) {
                    case "center" -> ParagraphAlignment.CENTER;
                    case "right" -> ParagraphAlignment.RIGHT;
                    case "justify" -> ParagraphAlignment.BOTH;
                    default -> ParagraphAlignment.LEFT;
                });
                this.blocks(element, context, format);
                context.close();
            }
        }
    }

    private void inline(Node node, Context context, Format format) {
        if (node instanceof TextNode text) {
            String value = text.text();
            if (context.paragraph == null && value.isBlank()) return;
            if (context.paragraph == null) value = value.stripLeading();
            if (value.isEmpty()) return;
            context.open();
            context.run(value, format);
            return;
        }
        if (!(node instanceof Element element)) return;
        Format inner = format.copy();
        switch (element.tagName()) {
            case "b", "strong" -> inner.bold = true;
            case "i", "em", "cite" -> inner.italic = true;
            case "u", "ins" -> inner.underline = true;
            case "s", "strike", "del" -> inner.strike = true;
            case "sub" -> inner.subscript = true;
            case "sup" -> inner.superscript = true;
            case "mark" -> inner.highlight = true;
            case "code", "kbd", "samp" -> inner.code = true;
            case "br" -> {
                context.open().createRun().addBreak();
                return;
            }
            case "img" -> {
                return;
            }
            case "a" -> {
                String href = element.attr("href");
                String label = element.text();
                if (!href.isBlank() && (href.startsWith("http") || href.startsWith("mailto:"))) {
                    DocxWriter.link(context.open(), href, label.isBlank() ? href : label);
                    return;
                }
            }
            default -> {
            }
        }
        this.blocks(element, context, inner);
    }

    private void table(Element table, Context context, Format format) {
        context.close();
        List<Element> rows = new ArrayList<>(table.select("tr"));
        rows.removeIf(row -> row.closest("table") != table);
        int columns = rows.stream().mapToInt(row -> row.select("> th, > td").size()).max().orElse(0);
        XWPFTable docTable = rows.isEmpty() || columns == 0 ? null : context.sink.table(rows.size(), columns);
        if (docTable == null) {
            // inside a table cell: one paragraph per row, cells separated by " | "
            for (Element row : rows) {
                List<String> cells = row.select("> th, > td").stream().map(Element::text).toList();
                context.open();
                context.run(String.join("  |  ", cells), format);
                context.close();
            }
            return;
        }
        boolean header = !rows.getFirst().select("> th").isEmpty();
        double[] widths = new double[columns];
        java.util.Arrays.fill(widths, 1);
        for (int r = 0; r < rows.size(); r++) {
            List<Element> cells = rows.get(r).select("> th, > td");
            for (int c = 0; c < columns; c++) {
                Sink.InCell cellSink = new Sink.InCell(docTable.getRow(r).getCell(c));
                if (c < cells.size()) {
                    new HtmlRenderer(this.writer).render(cells.get(c).html(), cellSink, header && r == 0 ? TemplateStyles.TABLE_HEADER : TemplateStyles.TABLE_TEXT);
                }
                if (docTable.getRow(r).getCell(c).getParagraphs().isEmpty()) docTable.getRow(r).getCell(c).addParagraph();
            }
        }
        this.writer.table(docTable, TemplateStyles.TABLE_GRID, widths, header);
    }

    private static String alignment(Element element) {
        String align = element.attr("align");
        if (align.isBlank()) {
            String style = element.attr("style").toLowerCase(Locale.ROOT);
            int index = style.indexOf("text-align");
            if (index >= 0) {
                String rest = style.substring(style.indexOf(':', index) + 1);
                align = rest.split(";")[0].trim();
            }
        }
        return align.isBlank() ? null : align.toLowerCase(Locale.ROOT);
    }

    private enum ListItem {BULLET, NUMBER}

    private static final class Format {
        boolean bold, italic, underline, strike, subscript, superscript, highlight, code;

        Format copy() {
            Format copy = new Format();
            copy.bold = this.bold;
            copy.italic = this.italic;
            copy.underline = this.underline;
            copy.strike = this.strike;
            copy.subscript = this.subscript;
            copy.superscript = this.superscript;
            copy.highlight = this.highlight;
            copy.code = this.code;
            return copy;
        }
    }

    /** The paragraph being filled and the list / quote nesting it belongs to. */
    private final class Context {
        final Sink sink;
        final String style;
        XWPFParagraph paragraph;
        int listLevel = -1;
        ListItem listItem;
        BigInteger listNumId;
        int quote;

        Context(Sink sink, String style) {
            this.sink = sink;
            this.style = style;
        }

        XWPFParagraph open() {
            if (this.paragraph != null) return this.paragraph;
            this.paragraph = DocxWriter.styled(this.sink.paragraph(), this.style);
            if (this.listItem == ListItem.BULLET) HtmlRenderer.this.writer.bullet(this.paragraph, this.listLevel);
            else if (this.listItem == ListItem.NUMBER) HtmlRenderer.this.writer.numbered(this.paragraph, this.listNumId, this.listLevel);
            if (this.listItem != null) this.paragraph.setSpacingAfter(40);
            if (this.quote > 0) this.paragraph.setIndentationLeft(360 * this.quote);
            // later paragraphs of the same list item are continuation paragraphs, not new bullets
            this.listItem = null;
            return this.paragraph;
        }

        void run(String text, Format format) {
            XWPFRun run = this.open().createRun();
            DocxWriter.setText(run, text);
            if (format.bold) run.setBold(true);
            if (format.italic) run.setItalic(true);
            if (format.underline) run.setUnderline(UnderlinePatterns.SINGLE);
            if (format.strike) run.setStrikeThrough(true);
            if (format.subscript) run.setVerticalAlignment(VerticalAlign.SUBSCRIPT.name().toLowerCase(Locale.ROOT));
            if (format.superscript) run.setVerticalAlignment(VerticalAlign.SUPERSCRIPT.name().toLowerCase(Locale.ROOT));
            if (format.highlight) run.setTextHighlightColor("yellow");
            if (format.code) run.setFontFamily("Consolas");
        }

        void close() {
            this.paragraph = null;
        }
    }
}
