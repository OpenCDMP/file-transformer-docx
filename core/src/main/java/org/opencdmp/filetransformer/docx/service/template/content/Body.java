package org.opencdmp.filetransformer.docx.service.template.content;

import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;

import java.util.List;

/**
 * The questionnaire content of a blueprint section or an asset, already resolved and pruned:
 * sub-sections, field sets, questions and their answers. Rendered by {@code {{>...body}}};
 * sub-section headings start at {@link #baseLevel} ({@code Heading<n>} styles of the template).
 */
public record Body(int baseLevel, List<Node> nodes) implements Content {

    public boolean isEmpty() {
        return this.nodes == null || this.nodes.isEmpty();
    }

    public sealed interface Node permits SectionNode, FieldSetNode {
    }

    public record SectionNode(String number, String title, RichText description, List<Node> children) implements Node {
    }

    /**
     * A field set (a group of questions). {@link #items} holds one entry per multiplicity item;
     * {@link #columns} holds the question labels when the set renders as a table.
     */
    public record FieldSetNode(String number, String title, RichText description, Mode mode, List<String> columns, List<Item> items, String comment) implements Node {
    }

    public enum Mode {
        /** Question label above the answer. */
        STACKED,
        /** Two-column question / answer table. */
        QUESTION_TABLE,
        /** Answers only. */
        COMPACT,
        /** Multiplicity table: one row per item, one column per question. */
        TABLE,
        /** Multiplicity table: one row per question, one column per item. */
        TRANSPOSED
    }

    public record Item(List<Question> questions) {
    }

    public record Question(String label, Answer answer) {
    }

    public sealed interface Answer permits Text, Html, Values, References, Image, Attachment, Missing {
    }

    public record Text(String value) implements Answer {
    }

    public record Html(String html) implements Answer {
    }

    /** Several values (multi-select, tags, ...), rendered as a list. */
    public record Values(List<String> values) implements Answer {
    }

    public record References(List<ReferenceView> references) implements Answer {
    }

    public record Image(byte[] data, String mimeType, String filename) implements Answer {
    }

    public record Attachment(String filename) implements Answer {
    }

    /** A visible question without an answer, kept only when the export shows unanswered questions. */
    public record Missing() implements Answer {
    }
}
