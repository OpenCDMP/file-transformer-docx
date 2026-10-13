package org.opencdmp.filetransformer.docx.service.template;

import org.opencdmp.commonmodels.enums.ReferenceTypeFieldDataType;
import org.opencdmp.commonmodels.models.reference.ReferenceFieldModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.commonmodels.models.referencetype.ReferenceTypeFieldModel;
import org.opencdmp.filetransformer.docx.model.PidLink;
import org.opencdmp.filetransformer.docx.service.pid.PidService;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.FactView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;

import java.time.Instant;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Turns a {@link ReferenceModel} into a {@link ReferenceView}: a clean label, a link when the
 * reference carries a known persistent identifier, and the fields selected for display with their
 * reference-type labels. Bookkeeping fields (source key, reference type, pid type) are never shown.
 */
final class ReferenceFormatter {

    private static final Set<String> BOOKKEEPING_FIELDS = Set.of("key", "referenceType", "pidTypeField", "source", "sourceType");
    private static final Set<String> LINK_FIELDS = Set.of("uri", "url", "link");
    /** The pseudo field selecting the persistent identifier (ORCID, DOI, ...) in the reference details setting. */
    static final String PID = "pid";
    // OpenAIRE-style labels carry the identifier: "Maria Papadopoulou (orcid:0000-0002-1825-0097)"
    private static final Pattern LABEL_WITH_PID = Pattern.compile("^(.*?)\\s*\\(([A-Za-z][A-Za-z0-9_-]*):\\s*([^()\\s]+)\\)\\s*$");

    private final PidService pidService;
    private final Map<String, List<String>> selectedFields;
    private final DateTimeFormatter dateFormatter;

    ReferenceFormatter(PidService pidService, Map<String, List<String>> selectedFields, DateTimeFormatter dateFormatter) {
        this.pidService = pidService;
        this.selectedFields = selectedFields;
        this.dateFormatter = dateFormatter;
    }

    ReferenceView format(ReferenceModel reference) {
        ReferenceView view = new ReferenceView();
        String label = firstNonBlank(reference.getLabel(), reference.getDescription(), reference.getReference());
        view.reference = reference.getReference();
        if (reference.getType() != null) {
            view.type = reference.getType().getName();
            view.typeCode = reference.getType().getCode();
        }

        Matcher matcher = label != null ? LABEL_WITH_PID.matcher(label) : null;
        if (matcher != null && matcher.matches() && !matcher.group(1).isBlank()) {
            label = matcher.group(1);
            view.pidType = matcher.group(2).toLowerCase(Locale.ROOT);
            view.pid = matcher.group(3);
        }
        view.label = label;

        List<ReferenceFieldModel> fields = reference.getDefinition() != null && reference.getDefinition().getFields() != null ? reference.getDefinition().getFields() : List.of();
        String linkFromFields = null;
        for (ReferenceFieldModel field : fields) {
            if (field.getCode() == null || isBlank(field.getValue())) continue;
            view.fields.put(Views.key(field.getCode()), field.getValue());
            if (field.getCode().equals("pidTypeField") && view.pidType == null) {
                view.pidType = field.getValue().toLowerCase(Locale.ROOT);
                view.pid = reference.getReference();
            }
            if (LINK_FIELDS.contains(field.getCode()) && linkFromFields == null) linkFromFields = field.getValue();
        }
        view.url = this.pidUrl(view.pidType, view.pid);
        if (view.url == null) view.url = linkFromFields;

        List<String> shown = this.fieldsToShow(reference, fields);
        // the identifier leads unless the configured fields place it ("pid") or leave it out
        if (!this.isSelected(reference) && !shown.contains(PID)) shown.addFirst(PID);
        for (String code : shown) {
            if (PID.equals(code)) {
                if (view.pid != null && view.pidType != null) view.details.add(new FactView(view.pidType.toUpperCase(Locale.ROOT), view.pid, this.pidUrl(view.pidType, view.pid)));
                continue;
            }
            ReferenceFieldModel field = fields.stream().filter(x -> code.equals(x.getCode())).findFirst().orElse(null);
            if (field == null || isBlank(field.getValue())) continue;
            view.details.add(new FactView(this.fieldLabel(reference, code), this.fieldValue(field), LINK_FIELDS.contains(code) ? field.getValue() : null));
        }
        List<String> parts = new ArrayList<>();
        for (FactView detail : view.details) parts.add(detail.label + ": " + detail.value);
        view.detailsText = String.join(" · ", parts);
        return view;
    }

    private boolean isSelected(ReferenceModel reference) {
        String typeCode = reference.getType() != null && reference.getType().getCode() != null ? reference.getType().getCode().toLowerCase(Locale.ROOT) : null;
        return typeCode != null && this.selectedFields.containsKey(typeCode);
    }

    private List<String> fieldsToShow(ReferenceModel reference, List<ReferenceFieldModel> fields) {
        String typeCode = reference.getType() != null && reference.getType().getCode() != null ? reference.getType().getCode().toLowerCase(Locale.ROOT) : null;
        if (typeCode != null && this.selectedFields.containsKey(typeCode)) return new ArrayList<>(this.selectedFields.get(typeCode));

        List<String> codes = new ArrayList<>();
        if (reference.getType() != null && reference.getType().getDefinition() != null && reference.getType().getDefinition().getFields() != null) {
            for (ReferenceTypeFieldModel typeField : reference.getType().getDefinition().getFields()) {
                if (typeField.getCode() != null) codes.add(typeField.getCode());
            }
        } else {
            for (ReferenceFieldModel field : fields) codes.add(field.getCode());
        }
        codes.removeIf(code -> code == null || BOOKKEEPING_FIELDS.contains(code) || isCatalogueId(code));
        return codes;
    }

    /**
     * A catalogue's own record id (e.g. FAIRsharing "internal-id"), which means nothing to a reader. The reference's
     * public identifier ({@code pid}) is shown separately; a type's details can still name such a field explicitly.
     */
    static boolean isCatalogueId(String code) {
        String normalized = code.toLowerCase(Locale.ROOT).replaceAll("[^a-z]", "");
        return !normalized.equals(PID) && (normalized.equals("id") || normalized.equals("internalid") || normalized.endsWith("internalid") || normalized.equals("recordid"));
    }

    private String fieldLabel(ReferenceModel reference, String code) {
        if (reference.getType() != null && reference.getType().getDefinition() != null && reference.getType().getDefinition().getFields() != null) {
            for (ReferenceTypeFieldModel typeField : reference.getType().getDefinition().getFields()) {
                if (code.equals(typeField.getCode()) && !isBlank(typeField.getLabel())) return typeField.getLabel();
            }
        }
        return Views.humanize(code);
    }

    private String fieldValue(ReferenceFieldModel field) {
        String value = field.getValue().trim();
        if (field.getDataType() == ReferenceTypeFieldDataType.Date || value.matches("\\d{4}-\\d{2}-\\d{2}(T.*)?")) {
            try {
                return value.length() == 10 ? this.dateFormatter.format(LocalDate.parse(value)) : this.dateFormatter.format(Instant.parse(value));
            } catch (Exception e) {
                return value;
            }
        }
        return value;
    }

    private String pidUrl(String pidType, String pid) {
        if (pidType == null || pid == null) return null;
        PidLink link = this.pidService.getPid(pidType);
        return link != null && link.getLink() != null ? link.getLink().replace("{pid}", pid) : null;
    }

    private static String firstNonBlank(String... values) {
        for (String value : values) if (!isBlank(value)) return value.trim();
        return null;
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }
}
