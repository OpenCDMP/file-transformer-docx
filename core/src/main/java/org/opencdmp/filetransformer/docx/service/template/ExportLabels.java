package org.opencdmp.filetransformer.docx.service.template;

import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 * Interface texts of an export in the plan's language ({@code export-labels/labels_<lang>.properties}),
 * falling back to English per key. Exposed to templates as {@code {{labels.<key>}}}.
 */
public final class ExportLabels {

    private static final String BUNDLE = "export-labels.labels";
    private static final ResourceBundle.Control NO_LOCALE_FALLBACK = ResourceBundle.Control.getNoFallbackControl(ResourceBundle.Control.FORMAT_PROPERTIES);

    private final Map<String, String> labels;

    private ExportLabels(Map<String, String> labels) {
        this.labels = labels;
    }

    public static ExportLabels forLanguage(String language) {
        Map<String, String> labels = new LinkedHashMap<>();
        ResourceBundle base = ResourceBundle.getBundle(BUNDLE, Locale.ROOT, NO_LOCALE_FALLBACK);
        for (String key : base.keySet()) labels.put(key, base.getString(key));
        if (language != null && !language.isBlank()) {
            try {
                ResourceBundle localized = ResourceBundle.getBundle(BUNDLE, Locale.forLanguageTag(language.trim()), NO_LOCALE_FALLBACK);
                for (String key : localized.keySet()) labels.put(key, localized.getString(key));
            } catch (MissingResourceException e) {
                // no translation for this language: English
            }
        }
        return new ExportLabels(labels);
    }

    /** These labels with some texts replaced; an empty replacement blanks the text. */
    public ExportLabels withOverrides(Map<String, String> overrides) {
        if (overrides == null || overrides.isEmpty()) return this;
        Map<String, String> labels = new LinkedHashMap<>(this.labels);
        labels.putAll(overrides);
        return new ExportLabels(labels);
    }

    public String get(String key) {
        return this.labels.getOrDefault(key, key);
    }

    public Map<String, String> asMap() {
        return this.labels;
    }
}
