package org.opencdmp.filetransformer.docx.service.template.content;

/** Rich text (HTML from the OpenCDMP editor), rendered with the template's paragraph and list styles. */
public record RichText(String html) implements Content {
}
