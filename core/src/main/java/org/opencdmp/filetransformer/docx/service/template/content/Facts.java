package org.opencdmp.filetransformer.docx.service.template.content;

import org.opencdmp.filetransformer.docx.service.template.view.ExportView.FactView;

import java.util.List;

/** Label / value pairs, rendered as a borderless key-value table (the {@code OpencdmpFactsTable} style). */
public record Facts(List<FactView> items) implements Content {
}
