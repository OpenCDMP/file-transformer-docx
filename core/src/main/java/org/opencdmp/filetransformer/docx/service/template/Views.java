package org.opencdmp.filetransformer.docx.service.template;

import org.jsoup.Jsoup;
import org.opencdmp.filetransformer.docx.service.template.content.Body;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/** Small helpers shared by the view building code. */
final class Views {

    private Views() {
    }

    /** A map key usable in template expressions: "rda.dataset.title" becomes "rda_dataset_title". */
    static String key(String raw) {
        return raw == null ? null : raw.trim().replaceAll("[^A-Za-z0-9]+", "_").replaceAll("^_+|_+$", "");
    }

    /** "startDate" / "start_date" become "Start date". */
    static String humanize(String code) {
        if (code == null || code.isBlank()) return code;
        String spaced = code.replaceAll("([a-z0-9])([A-Z])", "$1 $2").replaceAll("[_\\-.]+", " ").trim().toLowerCase(Locale.ROOT);
        return Character.toUpperCase(spaced.charAt(0)) + spaced.substring(1);
    }

    static boolean isBlankHtml(String html) {
        return html == null || html.isBlank() || (Jsoup.parse(html).text().isBlank() && !html.contains("<img"));
    }

    static String plainText(Body.Answer answer) {
        if (answer == null) return "";
        return switch (answer) {
            case Body.Text text -> text.value();
            case Body.Html html -> Jsoup.parse(html.html()).text();
            case Body.Values values -> String.join(", ", values.values());
            case Body.References references -> {
                List<String> labels = new ArrayList<>();
                references.references().forEach(x -> labels.add(x.label));
                yield String.join(", ", labels);
            }
            case Body.Image image -> image.filename() != null ? image.filename() : "";
            case Body.Attachment attachment -> attachment.filename() != null ? attachment.filename() : "";
            case Body.Missing missing -> "";
        };
    }

    /** Titles longer than this are set smaller on covers, so they wrap into fewer, fuller lines. */
    static final int LONG_TITLE = 38;

    static boolean isLongTitle(String title) {
        return title != null && title.strip().length() > LONG_TITLE;
    }

    static String join(String separator, String... parts) {
        List<String> present = new ArrayList<>();
        for (String part : parts) if (part != null && !part.isBlank()) present.add(part.trim());
        return String.join(separator, present);
    }
}
