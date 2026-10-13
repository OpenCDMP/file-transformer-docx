package org.opencdmp.filetransformer.docx.service.template.render;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.content.Facts;
import org.opencdmp.filetransformer.docx.service.template.content.PageBreak;
import org.opencdmp.filetransformer.docx.service.template.content.RichText;
import org.opencdmp.filetransformer.docx.service.template.content.TableOfContents;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordExportOptions;

import java.util.List;
import java.util.Map;

/**
 * Renders the {@code {{>name}}} tag: formatted content inserted in place of the paragraph holding
 * the tag. Text-like content (rich text, strings, references) takes that paragraph's style, so the
 * template decides how e.g. a cover description looks. A tag with no value removes its paragraph.
 */
public final class ContentRenderPolicy extends AbstractRenderPolicy<Object> {

    /** A table-of-contents entry pre-filled into the TOC field so it reads well before fields are updated. */
    public record TocEntry(int level, String text) {
    }

    private final Map<String, String> labels;
    private final WordExportOptions options;
    private final List<TocEntry> tocEntries;
    private DocxWriter writer;
    private BodyWriter bodyWriter;

    public ContentRenderPolicy(Map<String, String> labels, WordExportOptions options, List<TocEntry> tocEntries) {
        this.labels = labels;
        this.options = options;
        this.tocEntries = tocEntries;
    }

    @Override
    protected boolean validate(Object data) {
        // absent values are handled in doRender: they remove the paragraph instead of leaving it empty
        return true;
    }

    @Override
    public void doRender(RenderContext<Object> context) {
        XWPFRun run = context.getRun();
        XWPFParagraph placeholder = (XWPFParagraph) run.getParent();
        this.ensureWriters(context.getXWPFDocument());
        String style = placeholder.getStyle() != null ? placeholder.getStyle() : TemplateStyles.ANSWER;
        Sink sink = new Sink.BeforePlaceholder(BodyContainerFactory.getBodyContainer(run), run);

        Object data = context.getData();
        switch (data) {
            case null -> {
            }
            case RichText richText -> new HtmlRenderer(this.writer).render(richText.html(), sink, style);
            case Body body -> this.bodyWriter.body(body, sink);
            case Facts facts -> this.bodyWriter.facts(facts.items(), sink);
            case TableOfContents toc -> this.toc(sink);
            case PageBreak pageBreak -> DocxWriter.pageBreak(sink.paragraph());
            case List<?> list when !list.isEmpty() && list.getFirst() instanceof ReferenceView -> {
                @SuppressWarnings("unchecked") List<ReferenceView> references = (List<ReferenceView>) list;
                this.bodyWriter.references(references, sink, style);
            }
            default -> {
                String text = data.toString();
                if (!text.isBlank()) DocxWriter.text(DocxWriter.styled(sink.paragraph(), style), text, null);
            }
        }
        // drop the placeholder paragraph when the tag was all it held
        String tag = context.getTagSource() != null ? context.getTagSource().toString() : "";
        this.clearPlaceholder(context, placeholder.getText().trim().equals(tag.trim()));
    }

    private void ensureWriters(XWPFDocument document) {
        if (this.writer != null && this.writer.document() == document) return;
        this.writer = new DocxWriter(document);
        this.bodyWriter = new BodyWriter(this.writer, this.labels, this.options.getReferenceStyle(), this.options.getTableStyle());
    }

    /**
     * A TOC field whose current result already lists the document's chapters (without page numbers),
     * so it is readable as-is; the PDF conversion and Word (F9) replace it with the real table.
     */
    private void toc(Sink sink) {
        int depth = this.options.getTocDepth();
        String instruction = " TOC \\o \"1-" + depth + "\" \\h \\z \\u ";
        List<TocEntry> entries = this.tocEntries.stream().filter(x -> x.level() <= depth).toList();
        if (entries.isEmpty()) {
            XWPFParagraph paragraph = DocxWriter.styled(sink.paragraph(), TemplateStyles.TOC_PREFIX + 1);
            DocxWriter.field(paragraph, instruction, () -> DocxWriter.text(paragraph, "—", null));
        } else {
            XWPFParagraph first = null;
            XWPFParagraph last = null;
            for (TocEntry entry : entries) {
                XWPFParagraph paragraph = DocxWriter.styled(sink.paragraph(), TemplateStyles.TOC_PREFIX + Math.min(2, Math.max(1, entry.level())));
                if (first == null) {
                    first = paragraph;
                    paragraph.createRun().getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.BEGIN);
                    var instr = paragraph.createRun().getCTR().addNewInstrText();
                    instr.setStringValue(instruction);
                    instr.setSpace(org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute.Space.PRESERVE);
                    paragraph.createRun().getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.SEPARATE);
                }
                DocxWriter.text(paragraph, entry.text(), null);
                last = paragraph;
            }
            last.createRun().getCTR().addNewFldChar().setFldCharType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType.END);
        }
        // no w:updateFields: it makes Word prompt before showing the document; the PDF conversion
        // updates the table and Word users update it with F9
    }
}
