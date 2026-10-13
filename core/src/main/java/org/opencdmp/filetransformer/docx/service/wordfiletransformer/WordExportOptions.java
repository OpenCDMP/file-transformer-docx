package org.opencdmp.filetransformer.docx.service.wordfiletransformer;

import org.opencdmp.commonmodels.models.FileEnvelopeModel;

import java.time.format.DateTimeFormatter;

/**
 * Export options resolved from plugin configuration values. Branding options (accent color,
 * cover logo) are read from the definition/tenant scope (configuration fields); presentation
 * preferences (page size, date format, unanswered-question mode) are read from the exporting
 * user's plugin configuration (user configuration fields) and override nothing else — the two
 * sets are disjoint.
 */
public class WordExportOptions {

    public static final String FIELD_ACCENT_COLOR = "accent-color";
    public static final String FIELD_COVER_LOGO = "cover-logo";
    public static final String FIELD_PAGE_SIZE = "page-size";
    public static final String FIELD_DATE_FORMAT = "date-format";
    public static final String FIELD_UNANSWERED_QUESTIONS = "unanswered-questions";
    public static final String FIELD_INCLUDE_SECTIONS = "include-sections";
    public static final String FIELD_EXCLUDE_DESCRIPTIONS = "exclude-descriptions";

    // layout options (admin scope); honoured by '{{...}}' templates only
    public static final String FIELD_LAYOUT_ASSET_PLACEMENT = "layout-asset-placement";
    public static final String FIELD_LAYOUT_QUESTION_STYLE = "layout-question-style";
    public static final String FIELD_LAYOUT_REFERENCE_STYLE = "layout-reference-style";
    public static final String FIELD_LAYOUT_REFERENCE_FIELDS = "layout-reference-fields";
    public static final String FIELD_LAYOUT_TABLE_STYLE = "layout-table-style";
    public static final String FIELD_LAYOUT_COVER = "layout-cover";
    public static final String FIELD_LAYOUT_TOC = "layout-toc";
    public static final String FIELD_LAYOUT_TOC_DEPTH = "layout-toc-depth";
    public static final String FIELD_LAYOUT_NUMBERING = "layout-numbering";
    public static final String FIELD_LAYOUT_ASSET_PAGE_BREAK = "layout-asset-page-break";
    public static final String FIELD_LAYOUT_PLAN_FACTS = "layout-plan-facts";
    public static final String FIELD_LAYOUT_ASSET_FACTS = "layout-asset-facts";
    public static final String FIELD_LAYOUT_GUIDANCE = "layout-question-guidance";
    public static final String FIELD_LAYOUT_COMMENTS = "layout-comments";
    public static final String FIELD_LAYOUT_SECTION_PAGE_BREAK = "layout-section-page-break";
    public static final String FIELD_LAYOUT_LABELS = "layout-labels";
    public static final String FIELD_DRAFT_WATERMARK = "draft-watermark";
    public static final String FIELD_DRAFT_WATERMARK_TEXT = "draft-watermark-text";

    public static final String DEFAULT_DATE_FORMAT = "dd/MM/yyyy";

    public enum UnansweredQuestionsMode {HIDE, PLACEHOLDER}

    public enum PageSize {A4, LETTER}

    /** Where a plan's assets go: after the content of their blueprint section, or grouped before / after all sections. */
    public enum AssetPlacement {SECTIONS, BEFORE, AFTER}

    /** How a question is shown: label above its answer, a two-column question/answer table, or the answer alone. */
    public enum QuestionStyle {STACKED, TABLE, COMPACT}

    /** How reference answers are shown: label plus a details line, label only, or a table of the selected fields. */
    public enum ReferenceStyle {DETAILED, COMPACT, TABLE}

    /** Table look, mapped to a table style of the template (see {@code DocumentStyles}). */
    public enum TableStyle {LINED, GRID, STRIPED}

    /**
     * One row of a cover facts table: a key (a reference type code, a built-in key such as
     * {@code contacts} or {@code doi}, or {@code sem:<semantic>}) and an optional label override.
     */
    public record FactSpec(String key, String label) {
    }

    private String accentColor;
    private FileEnvelopeModel coverLogo;
    private PageSize pageSize;
    private String dateFormat = DEFAULT_DATE_FORMAT;
    private UnansweredQuestionsMode unansweredQuestions = UnansweredQuestionsMode.HIDE;
    private java.util.Set<String> includedSections;
    private java.util.Set<String> excludedDescriptions;
    private AssetPlacement assetPlacement = AssetPlacement.SECTIONS;
    private QuestionStyle questionStyle = QuestionStyle.STACKED;
    private ReferenceStyle referenceStyle = ReferenceStyle.DETAILED;
    private java.util.Map<String, java.util.List<String>> referenceFields = new java.util.HashMap<>();
    private TableStyle tableStyle = TableStyle.LINED;
    private boolean cover = true;
    private boolean toc = true;
    private int tocDepth = 2;
    private boolean numbering = true;
    private boolean assetPageBreak = false;
    private java.util.List<FactSpec> planFacts;
    private java.util.List<FactSpec> assetFacts;
    private boolean guidance = false;
    private boolean comments = true;
    private boolean sectionPageBreak = false;
    private java.util.Map<String, String> labelOverrides = new java.util.LinkedHashMap<>();
    private boolean draftWatermark = true;
    private String draftWatermarkText;

    public String getAccentColor() {
        return accentColor;
    }

    public void setAccentColor(String accentColor) {
        if (accentColor == null) return;
        String normalized = accentColor.trim().replace("#", "").toUpperCase();
        if (normalized.matches("[0-9A-F]{6}")) this.accentColor = normalized;
    }

    public FileEnvelopeModel getCoverLogo() {
        return coverLogo;
    }

    public void setCoverLogo(FileEnvelopeModel coverLogo) {
        this.coverLogo = coverLogo;
    }

    public PageSize getPageSize() {
        return pageSize;
    }

    public void setPageSize(String pageSize) {
        if (pageSize == null || pageSize.isBlank()) return;
        try {
            this.pageSize = PageSize.valueOf(pageSize.trim().toUpperCase());
        } catch (IllegalArgumentException e) {
            // unknown value: keep the template's page size
        }
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        if (dateFormat == null || dateFormat.isBlank()) return;
        try {
            DateTimeFormatter.ofPattern(dateFormat.trim());
            this.dateFormat = dateFormat.trim();
        } catch (IllegalArgumentException e) {
            // invalid pattern: keep the default
        }
    }

    public void setIncludedSections(String commaSeparatedIds) {
        this.includedSections = parseIds(commaSeparatedIds);
    }

    public void setExcludedDescriptions(String commaSeparatedIds) {
        this.excludedDescriptions = parseIds(commaSeparatedIds);
    }

    private java.util.Set<String> parseIds(String commaSeparatedIds) {
        if (commaSeparatedIds == null || commaSeparatedIds.isBlank()) return null;
        java.util.Set<String> ids = new java.util.HashSet<>();
        for (String id : commaSeparatedIds.split(",")) {
            if (!id.isBlank()) ids.add(id.trim().toLowerCase());
        }
        return ids.isEmpty() ? null : ids;
    }

    /** True when the section should be rendered; an absent include list means all sections. */
    public boolean isSectionIncluded(String sectionId) {
        if (this.includedSections == null) return true;
        return sectionId != null && this.includedSections.contains(sectionId.trim().toLowerCase());
    }

    public boolean isDescriptionExcluded(java.util.UUID descriptionId) {
        if (this.excludedDescriptions == null || descriptionId == null) return false;
        return this.excludedDescriptions.contains(descriptionId.toString().toLowerCase());
    }

    public UnansweredQuestionsMode getUnansweredQuestions() {
        return unansweredQuestions;
    }

    public void setUnansweredQuestions(String mode) {
        if (mode == null || mode.isBlank()) return;
        try {
            this.unansweredQuestions = UnansweredQuestionsMode.valueOf(mode.trim().toUpperCase());
        } catch (IllegalArgumentException e) {
            // unknown value: keep the default (hide)
        }
    }

    public AssetPlacement getAssetPlacement() {
        return assetPlacement;
    }

    public void setAssetPlacement(String value) {
        this.assetPlacement = parseEnum(AssetPlacement.class, value, this.assetPlacement);
    }

    public QuestionStyle getQuestionStyle() {
        return questionStyle;
    }

    public void setQuestionStyle(String value) {
        this.questionStyle = parseEnum(QuestionStyle.class, value, this.questionStyle);
    }

    public ReferenceStyle getReferenceStyle() {
        return referenceStyle;
    }

    public void setReferenceStyle(String value) {
        this.referenceStyle = parseEnum(ReferenceStyle.class, value, this.referenceStyle);
    }

    /**
     * Fields shown per reference type, keyed by lower-case reference type code; a type that is absent
     * shows every field its reference type defines. Format: {@code researchers:affiliation; grants:acronym,startDate}.
     */
    public java.util.Map<String, java.util.List<String>> getReferenceFields() {
        return referenceFields;
    }

    public void setReferenceFields(String value) {
        if (value == null || value.isBlank()) return;
        java.util.Map<String, java.util.List<String>> parsed = new java.util.HashMap<>();
        for (String entry : value.split(";")) {
            int colon = entry.indexOf(':');
            if (colon <= 0) continue;
            String type = entry.substring(0, colon).trim().toLowerCase();
            java.util.List<String> fields = new java.util.ArrayList<>();
            for (String field : entry.substring(colon + 1).split(",")) {
                if (!field.isBlank()) fields.add(field.trim());
            }
            if (!type.isEmpty()) parsed.put(type, fields);
        }
        this.referenceFields = parsed;
    }

    public TableStyle getTableStyle() {
        return tableStyle;
    }

    public void setTableStyle(String value) {
        this.tableStyle = parseEnum(TableStyle.class, value, this.tableStyle);
    }

    public boolean isCover() {
        return cover;
    }

    public void setCover(String value) {
        this.cover = parseBoolean(value, this.cover);
    }

    public boolean isToc() {
        return toc;
    }

    public void setToc(String value) {
        this.toc = parseBoolean(value, this.toc);
    }

    /** The heading levels the table of contents lists, from 1 (plan sections) to 6. */
    public int getTocDepth() {
        return tocDepth;
    }

    public void setTocDepth(String value) {
        if (value == null || value.isBlank()) return;
        try {
            this.tocDepth = Math.max(1, Math.min(6, Integer.parseInt(value.trim())));
        } catch (NumberFormatException e) {
            // invalid value: keep the default
        }
    }

    public boolean isNumbering() {
        return numbering;
    }

    public void setNumbering(String value) {
        this.numbering = parseBoolean(value, this.numbering);
    }

    public boolean isAssetPageBreak() {
        return assetPageBreak;
    }

    public void setAssetPageBreak(String value) {
        this.assetPageBreak = parseBoolean(value, this.assetPageBreak);
    }

    /** The rows of the plan cover facts, or {@code null} for the default rows. */
    public java.util.List<FactSpec> getPlanFacts() {
        return planFacts;
    }

    public void setPlanFacts(String value) {
        if (value == null || value.isBlank()) return;
        this.planFacts = parseFacts(value);
    }

    /** The rows of the asset cover facts, or {@code null} for the default rows. */
    public java.util.List<FactSpec> getAssetFacts() {
        return assetFacts;
    }

    public void setAssetFacts(String value) {
        if (value == null || value.isBlank()) return;
        this.assetFacts = parseFacts(value);
    }

    /** {@code "funders, grants=Project, doi"}: keys in order, each optionally relabelled. */
    private static java.util.List<FactSpec> parseFacts(String value) {
        if (value == null || value.isBlank()) return null;
        java.util.List<FactSpec> facts = new java.util.ArrayList<>();
        for (String entry : value.split("[,;]")) {
            if (entry.isBlank()) continue;
            int equals = entry.indexOf('=');
            String key = (equals >= 0 ? entry.substring(0, equals) : entry).trim();
            String label = equals >= 0 ? entry.substring(equals + 1).trim() : null;
            if (!key.isEmpty()) facts.add(new FactSpec(key, label == null || label.isEmpty() ? null : label));
        }
        return facts.isEmpty() ? null : facts;
    }

    /** Whether question and section guidance (the help texts of blueprints and asset templates) is exported. */
    public boolean isGuidance() {
        return guidance;
    }

    public void setGuidance(String value) {
        this.guidance = parseBoolean(value, this.guidance);
    }

    /** Whether field-set comments are exported. */
    public boolean isComments() {
        return comments;
    }

    public void setComments(String value) {
        this.comments = parseBoolean(value, this.comments);
    }

    public boolean isSectionPageBreak() {
        return sectionPageBreak;
    }

    public void setSectionPageBreak(String value) {
        this.sectionPageBreak = parseBoolean(value, this.sectionPageBreak);
    }

    /** Interface text overrides by label key, applied over the plan language's texts; an empty value blanks a text. */
    public java.util.Map<String, String> getLabelOverrides() {
        return labelOverrides;
    }

    /** {@code "asset=Dataset; generatedWith="}. */
    public void setLabelOverrides(String value) {
        if (value == null || value.isBlank()) return;
        java.util.Map<String, String> parsed = new java.util.LinkedHashMap<>();
        for (String entry : value.split(";")) {
            int equals = entry.indexOf('=');
            if (equals <= 0) continue;
            String key = entry.substring(0, equals).trim();
            if (!key.isEmpty()) parsed.put(key, entry.substring(equals + 1).trim());
        }
        this.labelOverrides = parsed;
    }

    /** Whether exports of plans / assets that are not finalized carry a watermark. */
    public boolean isDraftWatermark() {
        return draftWatermark;
    }

    public void setDraftWatermark(String value) {
        this.draftWatermark = parseBoolean(value, this.draftWatermark);
    }

    /** The watermark text, or {@code null} for the default ("DRAFT" in the plan's language). */
    public String getDraftWatermarkText() {
        return draftWatermarkText;
    }

    public void setDraftWatermarkText(String value) {
        this.draftWatermarkText = value == null || value.isBlank() ? null : value.trim();
    }

    private static <E extends Enum<E>> E parseEnum(Class<E> type, String value, E fallback) {
        if (value == null || value.isBlank()) return fallback;
        try {
            return Enum.valueOf(type, value.trim().toUpperCase());
        } catch (IllegalArgumentException e) {
            return fallback;
        }
    }

    private static boolean parseBoolean(String value, boolean fallback) {
        if (value == null || value.isBlank()) return fallback;
        return Boolean.parseBoolean(value.trim());
    }
}
