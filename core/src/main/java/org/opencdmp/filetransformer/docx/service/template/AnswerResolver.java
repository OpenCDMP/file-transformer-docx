package org.opencdmp.filetransformer.docx.service.template;

import org.opencdmp.commonmodels.enums.PlanAccessType;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.PlanUserModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.plan.PlanContactModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.FieldModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.RadioBoxDataModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SelectDataModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SystemFieldDataModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.UploadDataModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.service.language.LanguageService;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;

import java.math.BigDecimal;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

/**
 * Resolves the stored value of one question into a displayable {@link Body.Answer}: option values
 * become labels, dates are formatted, references are formatted, internal entries become the linked
 * plan / asset titles. Returns {@code null} for an unanswered question. Never throws for an
 * unsupported type: an unknown question type is treated as unanswered rather than failing the export.
 */
final class AnswerResolver {

    private static final Set<String> IMAGE_TYPES = Set.of("image/jpeg", "image/png", "image/gif", "image/bmp", "image/tiff");

    private final ExportLabels labels;
    private final DateTimeFormatter dateFormatter;
    private final ReferenceFormatter referenceFormatter;
    private final LanguageService languageService;
    private final FileStorageService fileStorageService;
    private final boolean useSharedStorage;
    private final PlanModel plan;
    private final Map<UUID, String> assetTitles;

    AnswerResolver(ExportLabels labels, DateTimeFormatter dateFormatter, ReferenceFormatter referenceFormatter, LanguageService languageService,
                   FileStorageService fileStorageService, boolean useSharedStorage, PlanModel plan, Map<UUID, String> assetTitles) {
        this.labels = labels;
        this.dateFormatter = dateFormatter;
        this.referenceFormatter = referenceFormatter;
        this.languageService = languageService;
        this.fileStorageService = fileStorageService;
        this.useSharedStorage = useSharedStorage;
        this.plan = plan;
        this.assetTitles = assetTitles;
    }

    /**
     * @param asset the asset whose question this is, or {@code null} for a blueprint question
     */
    Body.Answer resolve(FieldModel field, org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel value, DescriptionModel asset) {
        if (field == null || field.getData() == null || field.getData().getFieldType() == null) return null;
        return switch (field.getData().getFieldType()) {
            case SYSTEM -> this.system((SystemFieldDataModel) field.getData(), asset);
            case FREE_TEXT, TEXT_AREA -> value != null && !isBlank(value.getTextValue()) ? new Body.Text(value.getTextValue().trim()) : null;
            case RICH_TEXT_AREA -> value != null && !Views.isBlankHtml(value.getTextValue()) ? new Body.Html(value.getTextValue()) : null;
            case SELECT -> value != null ? this.select((SelectDataModel) field.getData(), value.getTextListValue()) : null;
            case RADIO_BOX -> value != null ? this.radio((RadioBoxDataModel) field.getData(), value.getTextValue()) : null;
            case BOOLEAN_DECISION -> value != null && value.getBooleanValue() != null ? new Body.Text(this.labels.get(value.getBooleanValue() ? "yes" : "no")) : null;
            case CHECK_BOX -> value != null && Boolean.TRUE.equals(value.getBooleanValue()) ? new Body.Text(this.labels.get("yes")) : null;
            case DATE_PICKER -> value != null && value.getDateValue() != null ? new Body.Text(this.dateFormatter.format(value.getDateValue())) : null;
            case NUMBER -> value != null && value.getNumberValue() != null ? new Body.Text(formatNumber(value.getNumberValue())) : null;
            case TAGS -> value != null ? values(value.getTextListValue()) : null;
            case REFERENCE_TYPES -> value != null ? this.references(value.getReferences()) : null;
            case DATASET_IDENTIFIER, VALIDATION -> value != null && value.getExternalIdentifier() != null && !isBlank(value.getExternalIdentifier().getIdentifier())
                    ? new Body.Text(Views.join(" ", value.getExternalIdentifier().getIdentifier(), isBlank(value.getExternalIdentifier().getType()) ? null : "(" + value.getExternalIdentifier().getType() + ")"))
                    : null;
            case UPLOAD -> value != null ? this.upload((UploadDataModel) field.getData(), value) : null;
            case INTERNAL_ENTRIES_DESCRIPTIONS -> value != null ? this.internalEntries(value.getTextListValue(), this.assetTitles) : null;
            case INTERNAL_ENTRIES_PlANS -> value != null ? this.internalEntries(value.getTextListValue(), this.plan != null && this.plan.getId() != null ? Map.of(this.plan.getId(), nonNull(this.plan.getLabel())) : Map.of()) : null;
        };
    }

    /** The question label: the field's label, the reference type name for an unlabelled reference question, or null. */
    static String label(FieldModel field) {
        if (field.getData() == null) return null;
        String label = field.getData().getLabel();
        if (isBlank(label) && field.getData() instanceof org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.ReferenceTypeDataModel reference
                && reference.getReferenceType() != null) {
            label = reference.getReferenceType().getName();
        }
        return isBlank(label) ? null : label.trim();
    }

    private Body.Answer system(SystemFieldDataModel data, DescriptionModel asset) {
        if (data.getSystemFieldType() == null) return null;
        return switch (data.getSystemFieldType()) {
            case PlanTitle -> text(this.plan != null ? this.plan.getLabel() : null);
            case PlanDescription -> this.plan != null && !Views.isBlankHtml(this.plan.getDescription()) ? new Body.Html(this.plan.getDescription()) : null;
            case PlanLanguage -> text(this.languageName(this.plan != null ? this.plan.getLanguage() : null));
            case PlanAccessRights -> this.plan != null && this.plan.getAccessType() != null ? new Body.Text(this.labels.get(this.plan.getAccessType() == PlanAccessType.Public ? "public" : "restricted")) : null;
            case PlanContact -> this.plan != null && this.plan.getProperties() != null ? values(contacts(this.plan.getProperties().getContacts())) : null;
            case PlanUser -> this.plan != null ? values(users(this.plan.getUsers())) : null;
            case DescriptionTitle -> text(asset != null ? asset.getLabel() : null);
            case DescriptionDescription -> asset != null && !Views.isBlankHtml(asset.getDescription()) ? new Body.Html(asset.getDescription()) : null;
            case DescriptionTag -> asset != null ? values(asset.getTags()) : null;
        };
    }

    private Body.Answer select(SelectDataModel data, List<String> selected) {
        if (selected == null || selected.isEmpty()) return null;
        List<String> labels = new ArrayList<>();
        for (String value : selected) {
            if (isBlank(value)) continue;
            String label = value;
            if (data.getOptions() != null) {
                for (SelectDataModel.OptionModel option : data.getOptions()) {
                    if (value.equals(option.getValue()) || value.equals(option.getLabel())) {
                        label = option.getLabel();
                        break;
                    }
                }
            }
            labels.add(label);
        }
        if (labels.isEmpty()) return null;
        return labels.size() == 1 && !Boolean.TRUE.equals(data.getMultipleSelect()) ? new Body.Text(labels.getFirst()) : new Body.Values(labels);
    }

    private Body.Answer radio(RadioBoxDataModel data, String selected) {
        if (isBlank(selected)) return null;
        if (data.getOptions() != null) {
            for (RadioBoxDataModel.RadioBoxOptionModel option : data.getOptions()) {
                if (selected.equals(option.getValue()) || selected.equals(option.getLabel())) return new Body.Text(option.getLabel());
            }
        }
        return new Body.Text(selected);
    }

    private Body.Answer references(List<ReferenceModel> references) {
        if (references == null || references.isEmpty()) return null;
        List<ReferenceView> views = new ArrayList<>();
        for (ReferenceModel reference : references) {
            if (reference == null) continue;
            ReferenceView view = this.referenceFormatter.format(reference);
            if (view.label != null) views.add(view);
        }
        return views.isEmpty() ? null : new Body.References(views);
    }

    private Body.Answer upload(UploadDataModel data, org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel value) {
        FileEnvelopeModel file = value.getFile();
        if (file == null || isBlank(value.getTextValue())) return null;
        if (file.getMimeType() != null && IMAGE_TYPES.contains(file.getMimeType())) {
            byte[] bytes = this.useSharedStorage && !isBlank(file.getFileRef()) ? this.fileStorageService.readFile(file.getFileRef()) : file.getFile();
            if (bytes != null && bytes.length > 0) return new Body.Image(bytes, file.getMimeType(), file.getFilename());
        }
        return isBlank(file.getFilename()) ? null : new Body.Attachment(file.getFilename());
    }

    private Body.Answer internalEntries(List<String> ids, Map<UUID, String> titles) {
        if (ids == null || ids.isEmpty()) return null;
        List<String> resolved = new ArrayList<>();
        for (String id : ids) {
            try {
                String title = titles.get(UUID.fromString(id.trim()));
                if (!isBlank(title)) resolved.add(title);
            } catch (IllegalArgumentException e) {
                // not an id: already a label
                if (!isBlank(id)) resolved.add(id.trim());
            }
        }
        return values(resolved);
    }

    private String languageName(String code) {
        if (isBlank(code)) return null;
        Language language = this.languageService.getLanguage(code);
        return language != null && !isBlank(language.getName()) ? language.getName() : code;
    }

    static List<String> contacts(List<PlanContactModel> contacts) {
        List<String> result = new ArrayList<>();
        if (contacts == null) return result;
        for (PlanContactModel contact : contacts) {
            if (contact == null) continue;
            String name = Views.join(" ", contact.getFirstName(), contact.getLastName());
            result.add(isBlank(contact.getEmail()) ? name : Views.join(" ", name, "(" + contact.getEmail() + ")"));
        }
        result.removeIf(String::isBlank);
        return result;
    }

    static List<String> users(List<PlanUserModel> users) {
        List<String> result = new ArrayList<>();
        if (users == null) return result;
        for (PlanUserModel user : users) {
            if (user == null || user.getUser() == null || isBlank(user.getUser().getName())) continue;
            result.add(user.getRole() != null ? user.getUser().getName() + " (" + Views.humanize(user.getRole().name()) + ")" : user.getUser().getName());
        }
        return result;
    }

    private static Body.Answer values(List<String> values) {
        if (values == null) return null;
        List<String> present = values.stream().filter(x -> !isBlank(x)).map(String::trim).toList();
        if (present.isEmpty()) return null;
        return present.size() == 1 ? new Body.Text(present.getFirst()) : new Body.Values(present);
    }

    private static Body.Answer text(String value) {
        return isBlank(value) ? null : new Body.Text(value.trim());
    }

    static String formatNumber(Double number) {
        BigDecimal decimal = BigDecimal.valueOf(number).stripTrailingZeros();
        return decimal.scale() <= 0 ? decimal.toBigInteger().toString() : decimal.toPlainString();
    }

    private static String nonNull(String value) {
        return value == null ? "" : value;
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }
}
