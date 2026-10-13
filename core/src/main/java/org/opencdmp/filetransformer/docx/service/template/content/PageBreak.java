package org.opencdmp.filetransformer.docx.service.template.content;

/** A page break, e.g. {@code {{?asset.newPage}}{{>pageBreak}}{{/asset.newPage}}}. */
public record PageBreak() implements Content {
}
