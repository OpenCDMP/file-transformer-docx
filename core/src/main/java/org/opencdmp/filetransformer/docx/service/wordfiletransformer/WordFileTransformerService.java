package org.opencdmp.filetransformer.docx.service.wordfiletransformer;

import gr.cite.tools.exception.MyApplicationException;
import gr.cite.tools.logging.LoggerService;
import org.apache.poi.xwpf.usermodel.*;
import org.opencdmp.commonmodels.enums.*;
import org.opencdmp.commonmodels.models.ConfigurationField;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.planblueprint.*;
import org.opencdmp.commonmodels.models.planreference.PlanReferenceModel;
import org.opencdmp.commonmodels.models.plugin.PluginFieldModel;
import org.opencdmp.commonmodels.models.plugin.PluginModel;
import org.opencdmp.commonmodels.models.plugin.PluginUserFieldModel;
import org.opencdmp.commonmodels.models.plugin.PluginUserModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerClient;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerConfiguration;
import org.opencdmp.filetransformer.docx.model.enums.FileFormats;
import org.opencdmp.filetransformerbase.models.misc.*;
import org.opencdmp.filetransformer.docx.service.pdf.PdfService;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.template.ExportLabels;
import org.opencdmp.filetransformer.docx.service.template.TemplateRenderer;
import org.opencdmp.filetransformer.docx.service.template.TemplateTexts;
import org.opencdmp.filetransformer.docx.service.template.render.TemplateStyles;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.validation.DocxTemplateValidationService;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.word.DocumentStyles;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.word.DraftWatermarkHelper;
import org.apache.poi.ooxml.POIXMLProperties;
import org.jsoup.Jsoup;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Component;
import org.springframework.web.context.annotation.RequestScope;

import java.io.*;
import java.util.*;
import java.util.List;

@Component
@RequestScope
public class WordFileTransformerService implements FileTransformerClient {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(WordFileTransformerService.class));

    private final static List<FileFormat> FILE_FORMATS = List.of(
            new FileFormat(FileFormats.PDF.getValue(), true, "fa-file-pdf-o"),
            new FileFormat(FileFormats.DOCX.getValue(), true, "fa-file-word-o"));

    private final static List<PluginEntityType> FILE_TRANSFORMER_ENTITY_TYPES = List.of(
            PluginEntityType.Plan, PluginEntityType.Description);

    private static final String FIELD_PLAN_TEMPLATE = "plan-docx-template-id";
    private static final String FIELD_DESCRIPTION_TEMPLATE = "description-docx-template-id";
    private static final String MIME_TYPE_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private static final String MIME_TYPE_PDF = "application/pdf";

    private final WordFileTransformerServiceProperties wordFileTransformerServiceProperties;
    private final PdfService pdfService;
    private final FileStorageService storageService;
    private final ResourceLoader resourceLoader;
    private final DocxTemplateValidationService templateValidationService;
    private final TemplateRenderer templateRenderer;
    @Autowired
    public WordFileTransformerService(
            WordFileTransformerServiceProperties wordFileTransformerServiceProperties,
            PdfService pdfService, FileStorageService storageService, ResourceLoader resourceLoader,
            DocxTemplateValidationService templateValidationService, TemplateRenderer templateRenderer) {
        this.wordFileTransformerServiceProperties = wordFileTransformerServiceProperties;
	    this.pdfService = pdfService;
	    this.storageService = storageService;
        this.resourceLoader = resourceLoader;
        this.templateValidationService = templateValidationService;
        this.templateRenderer = templateRenderer;
    }

    @Override
    public FileEnvelopeModel exportPlan(PlanModel plan, String variant) {
        FileFormats fileFormat = FileFormats.of(variant);
        byte[] bytes = this.buildPlanWordDocument(plan);
        String filename = switch (fileFormat) {
	        case DOCX -> this.getPlanFileName(plan, ".docx");
	        case PDF -> {
		        bytes = this.pdfService.convertToPDF(bytes);
		        yield this.getPlanFileName(plan, ".pdf");
	        }
	        default -> throw new MyApplicationException("Invalid type " + fileFormat);
        };
        
        FileEnvelopeModel wordFile = new FileEnvelopeModel();
        if (this.getConfiguration().isUseSharedStorage()) {
            String fileRef = this.storageService.storeFile(bytes);
            wordFile.setFileRef(fileRef);
        } else {
            wordFile.setFile(bytes);
        }
        wordFile.setFilename(filename);
        wordFile.setMimeType(fileFormat == FileFormats.PDF ? MIME_TYPE_PDF : MIME_TYPE_DOCX);
        return wordFile;
    }

    @Override
    public FileEnvelopeModel exportDescription(DescriptionModel descriptionModel, String variant) {
        FileFormats fileFormat = FileFormats.of(variant);
        byte[] bytes = this.buildDescriptionWordDocument(descriptionModel);
        String filename = switch (fileFormat) {
            case DOCX -> this.getDescriptionFileName(descriptionModel, ".docx");
            case PDF -> {
                bytes = this.pdfService.convertToPDF(bytes);
                yield this.getDescriptionFileName(descriptionModel, ".pdf");
            }
            default -> throw new MyApplicationException("Invalid type " + fileFormat);
        };

        FileEnvelopeModel wordFile = new FileEnvelopeModel();
        if (this.getConfiguration().isUseSharedStorage()) {
            String fileRef = this.storageService.storeFile(bytes);
            wordFile.setFileRef(fileRef);
        } else {
            wordFile.setFile(bytes);
        }
        wordFile.setFilename(filename);
        wordFile.setMimeType(fileFormat == FileFormats.PDF ? MIME_TYPE_PDF : MIME_TYPE_DOCX);
        return wordFile;
    }


    @Override
    public PlanModel importPlan(PlanImportModel planImportModel) {
        throw new MyApplicationException("import not supported");
    }

    @Override
    public DescriptionModel importDescription(DescriptionImportModel descriptionImportModel) {
        throw new MyApplicationException("import not supported");
    }

    @Override
    public FileTransformerConfiguration getConfiguration() {
        FileTransformerConfiguration configuration = new FileTransformerConfiguration();
        configuration.setFileTransformerId(this.wordFileTransformerServiceProperties.getTransformerId());
        configuration.setExportVariants(FILE_FORMATS);
        configuration.setImportVariants(null);
        configuration.setExportEntityTypes(FILE_TRANSFORMER_ENTITY_TYPES);
        configuration.setUseSharedStorage(this.wordFileTransformerServiceProperties.isUseSharedStorage());
        configuration.setConfigurationFields(this.wordFileTransformerServiceProperties.getConfigurationFields());
        configuration.setUserConfigurationFields(this.wordFileTransformerServiceProperties.getUserConfigurationFields());
        return configuration;
    }

    @Override
    public org.opencdmp.filetransformerbase.models.misc.TemplateValidationResultModel validateTemplate(FileEnvelopeModel fileEnvelopeModel, PluginEntityType entityType) {
        if (fileEnvelopeModel == null) throw new MyApplicationException("file required");
        // only DOCX files are templates; other file fields (e.g. the cover logo image) are not validated
        if (fileEnvelopeModel.getMimeType() != null && !fileEnvelopeModel.getMimeType().isBlank()
                && !fileEnvelopeModel.getMimeType().equals(MIME_TYPE_DOCX)) {
            return null;
        }
        byte[] bytes;
        if (this.wordFileTransformerServiceProperties.isUseSharedStorage() && fileEnvelopeModel.getFileRef() != null && !fileEnvelopeModel.getFileRef().isBlank()) {
            bytes = this.storageService.readFile(fileEnvelopeModel.getFileRef());
        } else {
            bytes = fileEnvelopeModel.getFile();
        }
        if (bytes == null) throw new MyApplicationException("file content required");
        return this.templateValidationService.validate(bytes, entityType != null ? entityType : PluginEntityType.Plan);
    }

    @Override
    public PreprocessingPlanModel preprocessingPlan(FileEnvelopeModel fileEnvelopeModel) {
        throw new MyApplicationException("preprocessing not supported");
    }

    @Override
    public PreprocessingDescriptionModel preprocessingDescription(FileEnvelopeModel fileEnvelopeModel) {
        throw new MyApplicationException("preprocessing not supported");
    }

    private List<ReferenceModel> getReferenceModelOfTypeCode(PlanModel plan, String code, UUID blueprintId){
        List<ReferenceModel> response = new ArrayList<>();
        if (plan.getReferences() == null) return response;
        for (PlanReferenceModel planReferenceModel : plan.getReferences()){
            if (planReferenceModel.getReference() != null && planReferenceModel.getReference().getType() != null && planReferenceModel.getReference().getType().getCode() != null  && planReferenceModel.getReference().getType().getCode().equals(code)){
                if (blueprintId == null || (planReferenceModel.getData() != null && blueprintId.equals(planReferenceModel.getData().getBlueprintFieldId()))) response.add(planReferenceModel.getReference());
                
            }
        }
        return response;
    }
    

    private byte[] buildPlanWordDocument(PlanModel planEntity) {
        if (planEntity == null) throw new MyApplicationException("planEntity required");
        PlanBlueprintModel planBlueprintModel = planEntity.getPlanBlueprint();
        if (planBlueprintModel == null) throw new MyApplicationException("PlanBlueprint required");
        if (planBlueprintModel.getDefinition() == null) throw new MyApplicationException("PlanBlueprint Definition required");
        if (planBlueprintModel.getDefinition().getSections() == null) throw new MyApplicationException("PlanBlueprint Section required");

        XWPFDocument document = this.loadTemplate(planBlueprintModel.getDefinition().getPlugins(), PluginEntityType.Plan, this.wordFileTransformerServiceProperties.getWordPlanTemplate());
        WordExportOptions exportOptions = this.resolveOptions(planBlueprintModel.getDefinition().getPlugins(), planEntity);
        document = this.templateRenderer.renderPlan(document, planEntity, exportOptions);
        this.finishTemplateDocument(document, exportOptions, planEntity, null);
        return this.toBytes(document);
    }

    private String getPlanFileName(PlanModel planModel, String extension){
        if (planModel == null) throw new MyApplicationException("PlanEntity required");

        List<ReferenceModel> grants = this.getReferenceModelOfTypeCode(planModel, this.wordFileTransformerServiceProperties.getGrantReferenceCode(), null);
        String fileName = null;
        if (planModel.getLabel() != null && !planModel.getLabel().isBlank()){
            fileName = planModel.getLabel();
        } else if (!grants.isEmpty() && grants.getFirst().getLabel() != null) {
            fileName = "PLAN_" + grants.getFirst().getLabel();
            fileName += "_" + planModel.getVersion();

        }

        return this.safeFileName(fileName, "plan") + extension;
    }

    /** Keeps letters of any script; replaces what file systems reject. */
    private String safeFileName(String name, String fallback) {
        String safe = name == null ? "" : name.replaceAll("[\\\\/:*?\"<>|\\p{Cntrl}]", "_").replaceAll("[\\s.]+$", "").strip();
        return safe.isBlank() ? fallback : safe;
    }

    private byte[] buildDescriptionWordDocument(DescriptionModel descriptionModel) {
        if (descriptionModel == null) throw new MyApplicationException("DescriptionEntity required");
        PlanModel planEntity = descriptionModel.getPlan();
        if (planEntity == null)  throw new MyApplicationException("plan is invalid");

        List<PluginModel> plugins = descriptionModel.getDescriptionTemplate() != null && descriptionModel.getDescriptionTemplate().getDefinition() != null
                ? descriptionModel.getDescriptionTemplate().getDefinition().getPlugins()
                : null;
        XWPFDocument document = this.loadTemplate(plugins, PluginEntityType.Description, this.wordFileTransformerServiceProperties.getWordDescriptionTemplate());
        WordExportOptions exportOptions = this.resolveOptions(plugins, planEntity);
        document = this.templateRenderer.renderAsset(document, descriptionModel, exportOptions);
        this.finishTemplateDocument(document, exportOptions, planEntity, descriptionModel);
        return this.toBytes(document);
    }

    private String getDescriptionFileName(DescriptionModel descriptionModel, String extension){
        if (descriptionModel == null) throw new MyApplicationException("DescriptionEntity required");
        return this.safeFileName(descriptionModel.getLabel(), "description") + extension;
    }

    /**
     * Resolves the export options: the defaults the plugin declares for its fields first, then the
     * definition / tenant scope (configuration fields), then the exporting user's scope (user
     * configuration fields). Blank values never override: they fall through to the previous layer.
     */
    private WordExportOptions resolveOptions(List<PluginModel> definitionPlugins, PlanModel plan) {
        WordExportOptions options = new WordExportOptions();
        String transformerId = this.wordFileTransformerServiceProperties.getTransformerId();

        // declared defaults: what editors show for an empty field
        if (this.wordFileTransformerServiceProperties.getConfigurationFields() != null) {
            for (ConfigurationField field : this.wordFileTransformerServiceProperties.getConfigurationFields()) {
                if (field.getCode() != null && field.getDefaultValue() != null) this.applyConfigurationField(options, field.getCode(), field.getDefaultValue(), null);
            }
        }
        if (this.wordFileTransformerServiceProperties.getUserConfigurationFields() != null) {
            for (ConfigurationField field : this.wordFileTransformerServiceProperties.getUserConfigurationFields()) {
                if (field.getCode() != null && field.getDefaultValue() != null) this.applyUserField(options, field.getCode(), field.getDefaultValue());
            }
        }

        // branding and layout: tenant / blueprint / asset-template scope (configuration fields)
        if (definitionPlugins != null) {
            definitionPlugins.stream()
                    .filter(x -> transformerId.equals(x.getCode()) && PluginType.FileTransformer.equals(x.getType()) && x.getFields() != null)
                    .findFirst()
                    .ifPresent(plugin -> {
                        for (PluginFieldModel field : plugin.getFields()) {
                            if (field.getCode() != null) this.applyConfigurationField(options, field.getCode(), field.getTextValue(), field.getFile());
                        }
                    });
        }

        // per-export choices (sections and assets to include), sent as user configuration fields
        if (plan != null && plan.getCreator() != null && plan.getCreator().getAdditionalInfo() != null && plan.getCreator().getAdditionalInfo().getPluginConfigurations() != null) {
            plan.getCreator().getAdditionalInfo().getPluginConfigurations().stream()
                    .filter(x -> transformerId.equals(x.getCode()) && PluginType.FileTransformer.equals(x.getType()) && x.getFields() != null)
                    .findFirst()
                    .ifPresent(plugin -> {
                        for (PluginUserFieldModel field : plugin.getFields()) {
                            if (field.getCode() != null) this.applyUserField(options, field.getCode(), field.getTextValue());
                        }
                    });
        }

        return options;
    }

    private void applyConfigurationField(WordExportOptions options, String code, String textValue, FileEnvelopeModel file) {
        switch (code) {
            case WordExportOptions.FIELD_ACCENT_COLOR -> options.setAccentColor(textValue);
            case WordExportOptions.FIELD_COVER_LOGO -> options.setCoverLogo(file);
            case WordExportOptions.FIELD_LAYOUT_ASSET_PLACEMENT -> options.setAssetPlacement(textValue);
            case WordExportOptions.FIELD_LAYOUT_QUESTION_STYLE -> options.setQuestionStyle(textValue);
            case WordExportOptions.FIELD_LAYOUT_REFERENCE_STYLE -> options.setReferenceStyle(textValue);
            case WordExportOptions.FIELD_LAYOUT_REFERENCE_FIELDS -> options.setReferenceFields(textValue);
            case WordExportOptions.FIELD_LAYOUT_TABLE_STYLE -> options.setTableStyle(textValue);
            case WordExportOptions.FIELD_LAYOUT_COVER -> options.setCover(textValue);
            case WordExportOptions.FIELD_LAYOUT_TOC -> options.setToc(textValue);
            case WordExportOptions.FIELD_LAYOUT_TOC_DEPTH -> options.setTocDepth(textValue);
            case WordExportOptions.FIELD_LAYOUT_NUMBERING -> options.setNumbering(textValue);
            case WordExportOptions.FIELD_LAYOUT_ASSET_PAGE_BREAK -> options.setAssetPageBreak(textValue);
            case WordExportOptions.FIELD_LAYOUT_PLAN_FACTS -> options.setPlanFacts(textValue);
            case WordExportOptions.FIELD_LAYOUT_ASSET_FACTS -> options.setAssetFacts(textValue);
            case WordExportOptions.FIELD_LAYOUT_GUIDANCE -> options.setGuidance(textValue);
            case WordExportOptions.FIELD_LAYOUT_COMMENTS -> options.setComments(textValue);
            case WordExportOptions.FIELD_LAYOUT_SECTION_PAGE_BREAK -> options.setSectionPageBreak(textValue);
            case WordExportOptions.FIELD_LAYOUT_LABELS -> options.setLabelOverrides(textValue);
            case WordExportOptions.FIELD_DRAFT_WATERMARK -> options.setDraftWatermark(textValue);
            case WordExportOptions.FIELD_DRAFT_WATERMARK_TEXT -> options.setDraftWatermarkText(textValue);
            case WordExportOptions.FIELD_PAGE_SIZE -> options.setPageSize(textValue);
            case WordExportOptions.FIELD_DATE_FORMAT -> options.setDateFormat(textValue);
            case WordExportOptions.FIELD_UNANSWERED_QUESTIONS -> options.setUnansweredQuestions(textValue);
            default -> { }
        }
    }

    /**
     * Per-export choices only: which sections and assets to include. Page size, date format and unanswered
     * questions are organisation settings now, so values users stored for them earlier are ignored rather than
     * overriding their organisation's.
     */
    private void applyUserField(WordExportOptions options, String code, String textValue) {
        switch (code) {
            case WordExportOptions.FIELD_INCLUDE_SECTIONS -> options.setIncludedSections(textValue);
            case WordExportOptions.FIELD_EXCLUDE_DESCRIPTIONS -> options.setExcludedDescriptions(textValue);
            default -> { }
        }
    }

    /** Post-processing after rendering: branding colour, page size, core properties, draft watermark. */
    private void finishTemplateDocument(XWPFDocument document, WordExportOptions options, PlanModel plan, DescriptionModel description) {
        DocumentStyles.recolor(document, TemplateStyles.ACCENT_COLOR, options.getAccentColor());
        if (options.getPageSize() != null) DocumentStyles.applyPageSize(document, options.getPageSize() == WordExportOptions.PageSize.A4);
        this.applyDocumentProperties(document, plan, description);
        if (options.isDraftWatermark() && !this.isFinalized(plan, description)) {
            String text = options.getDraftWatermarkText() != null ? options.getDraftWatermarkText() : this.watermarkText(plan, description, options);
            DraftWatermarkHelper.applyWatermark(document, text, TemplateStyles.WATERMARK_FONT, TemplateStyles.WATERMARK_COLOR, "1",
                    TemplateStyles.WATERMARK_WIDTH, TemplateStyles.WATERMARK_HEIGHT);
        }
    }

    private byte[] toBytes(XWPFDocument document) {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            document.write(out);
            return out.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void applyDocumentProperties(XWPFDocument document, PlanModel plan, DescriptionModel description) {
        try {
            POIXMLProperties.CoreProperties core = document.getProperties().getCoreProperties();
            String title = description != null ? description.getLabel() : plan.getLabel();
            if (title != null && !title.isBlank()) core.setTitle(title);
            if (plan.getCreator() != null && plan.getCreator().getName() != null && !plan.getCreator().getName().isBlank()) core.setCreator(plan.getCreator().getName());
            String rawDescription = description != null ? description.getDescription() : plan.getDescription();
            if (rawDescription != null && !rawDescription.isBlank()) core.setDescription(Jsoup.parse(rawDescription).text());
            if (description != null && description.getTags() != null && !description.getTags().isEmpty()) core.setKeywords(String.join(", ", description.getTags()));
            if (plan.getLanguage() != null && !plan.getLanguage().isBlank()) core.getUnderlyingProperties().setLanguageProperty(plan.getLanguage());
        } catch (Exception e) {
            logger.error("could not apply document core properties", e);
        }
    }

    /**
     * What a document that is not finalized is stamped with, by default: its status ("UNDER REVIEW"), which says
     * more than a fixed DRAFT; the DRAFT label, in the plan's language, when the status has no name.
     */
    private String watermarkText(PlanModel plan, DescriptionModel description, WordExportOptions options) {
        String status = description != null
                ? (description.getStatus() != null ? description.getStatus().getName() : null)
                : (plan.getStatus() != null ? plan.getStatus().getName() : null);
        java.util.Locale locale = plan.getLanguage() != null && !plan.getLanguage().isBlank() ? java.util.Locale.forLanguageTag(plan.getLanguage()) : java.util.Locale.ROOT;
        if (status != null && !status.isBlank()) return status.strip().toUpperCase(locale);
        return ExportLabels.forLanguage(plan.getLanguage()).withOverrides(options.getLabelOverrides()).get("draft");
    }

    private boolean isFinalized(PlanModel plan, DescriptionModel description) {
        if (description != null) return description.getStatus() != null && DescriptionStatus.Finalized.equals(description.getStatus().getInternalStatus());
        return plan.getStatus() != null && PlanStatus.Finalized.equals(plan.getStatus().getInternalStatus());
    }

    /** The custom template configured for the entity, or else the built-in one. */
    private XWPFDocument loadTemplate(List<PluginModel> plugins, PluginEntityType entityType, String defaultTemplate) {
        XWPFDocument document = plugins != null && !plugins.isEmpty() ? this.getCustomDocument(plugins, entityType) : null;
        if (document != null) return document;
        try {
            Resource resource = this.resourceLoader.getResource(defaultTemplate);
            try (InputStream inputStream = resource.getInputStream()) {
                return new XWPFDocument(inputStream);
            }
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
            throw new RuntimeException(e);
        }
    }

    private XWPFDocument getCustomDocument(List<PluginModel> plugins, PluginEntityType entityType) {
        try {
            if (plugins != null && !plugins.isEmpty()) {
                PluginModel plugin = plugins.stream().filter(x -> x.getCode().equals(this.wordFileTransformerServiceProperties.getTransformerId()) && x.getType().equals(PluginType.FileTransformer)).findFirst().orElse(null);
                if (plugin != null && plugin.getFields() != null) {
                    String templateCode = entityType == PluginEntityType.Description ? FIELD_DESCRIPTION_TEMPLATE : FIELD_PLAN_TEMPLATE;
                    PluginFieldModel field = plugin.getFields().stream().filter(y -> templateCode.equals(y.getCode())).findFirst().orElse(null);
                    if (field != null && field.getFile() != null && field.getFile().getFile() != null) {
                        XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(field.getFile().getFile()));
                        if (TemplateTexts.hasOldPlaceholders(document)) {
                            logger.warn("custom template uses the unsupported '{OPENCDMP.X}' placeholders.. fallback to default. Entity type: " + entityType);
                            return null;
                        }
                        return document;
                    }
                }
            }
        } catch (Exception e) {
            logger.error(e.getMessage());
            logger.warn("error creating custom document.. fallback to default. Entity type: " + entityType);
        }
        return null;
    }
}
