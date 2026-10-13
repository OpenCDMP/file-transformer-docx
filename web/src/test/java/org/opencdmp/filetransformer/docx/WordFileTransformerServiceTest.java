package org.opencdmp.filetransformer.docx;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.opencdmp.commonmodels.enums.DataType;
import org.opencdmp.commonmodels.enums.DescriptionStatus;
import org.opencdmp.commonmodels.enums.FieldType;
import org.opencdmp.commonmodels.enums.PlanAccessType;
import org.opencdmp.commonmodels.enums.PlanStatus;
import org.opencdmp.commonmodels.enums.PluginEntityType;
import org.opencdmp.commonmodels.enums.PluginType;
import org.opencdmp.commonmodels.models.ConfigurationField;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.description.DescriptionStatusModel;
import org.opencdmp.commonmodels.models.description.PropertyDefinitionModel;
import org.opencdmp.commonmodels.models.descriptiotemplate.DescriptionTemplateModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.plan.PlanPropertiesModel;
import org.opencdmp.commonmodels.models.planblueprint.PlanBlueprintModel;
import org.opencdmp.commonmodels.models.plan.PlanStatusModel;
import org.opencdmp.commonmodels.models.planblueprint.DefinitionModel;
import org.opencdmp.commonmodels.models.planblueprint.RootSectionModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.FieldModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.FieldSetModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.SectionModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.LabelDataModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.PropertyDefinitionFieldSetItemModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.PropertyDefinitionFieldSetModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.VisibilityStateModel;
import org.opencdmp.commonmodels.models.plugin.PluginFieldModel;
import org.opencdmp.commonmodels.models.plugin.PluginModel;
import org.opencdmp.commonmodels.models.plugin.PluginUserFieldModel;
import org.opencdmp.commonmodels.models.plugin.PluginUserModel;
import org.opencdmp.commonmodels.models.user.UserAdditionalInfoModel;
import org.opencdmp.commonmodels.models.user.UserModel;
import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.model.PidLink;
import org.opencdmp.filetransformer.docx.service.language.LanguageService;
import org.opencdmp.filetransformer.docx.service.pdf.PdfService;
import org.opencdmp.filetransformer.docx.service.pid.PidService;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.template.ExportViewBuilder;
import org.opencdmp.filetransformer.docx.service.template.TemplateRenderer;
import org.opencdmp.filetransformer.docx.service.template.render.TemplateStyles;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordExportOptions;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerService;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.springframework.core.io.DefaultResourceLoader;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Golden-path regression tests for the DOCX generation from the built-in and custom '{{...}}'
 * templates: structure, layout options, answers, branding, user preferences, draft watermark,
 * document core properties and validation, and the fallback for templates made with the
 * unsupported '{OPENCDMP.X}' placeholders. No Spring context and no external services required.
 */
class WordFileTransformerServiceTest {

    private static final String SECTION_ID = "sec-1";
    private static final String FIELDSET_ID = "fs-1";
    private static final String TEXT_FIELD_ID = "f-text";
    private static final String RICH_FIELD_ID = "f-rich";
    private static final String EMPTY_FIELDSET_ID = "fs-2";
    private static final String EMPTY_FIELD_ID = "f-empty";

    private WordFileTransformerServiceProperties properties;

    @BeforeEach
    void setup() {
        properties = new WordFileTransformerServiceProperties();
        properties.setTransformerId("docx-file-transformer");
        properties.setUseSharedStorage(false);
        properties.setWordPlanTemplate("classpath:documents/plan.docx");
        properties.setWordDescriptionTemplate("classpath:documents/description.docx");
        properties.setOrganizationReferenceCode("organisations");
        properties.setGrantReferenceCode("grants");
        properties.setResearcherReferenceCode("researchers");
        properties.setLicenceReferenceCode("licenses");
        properties.setDatasetReferenceCode("datasets");
        properties.setPublicationReferenceCode("publications");
        properties.setConfigurationFields(List.of(
                new ConfigurationField("plan-docx-template-id", DataType.File, "Plan Docx Template", List.of(PluginEntityType.Plan), false),
                new ConfigurationField("description-docx-template-id", DataType.File, "Description Docx Template", List.of(PluginEntityType.Description), false)));
    }

    private WordFileTransformerService service() {
        FileStorageService storage = new FileStorageService() {
            @Override
            public String storeFile(byte[] data) {
                throw new UnsupportedOperationException("shared storage disabled in tests");
            }

            @Override
            public byte[] readFile(String fileRef) {
                throw new UnsupportedOperationException("shared storage disabled in tests");
            }
        };
        PidService pidService = new PidService() {
            @Override
            public PidLink getPid(String pidType) {
                return null;
            }

            @Override
            public List<PidLink> loadPidLinks() {
                return List.of();
            }
        };
        LanguageService languageService = new LanguageService() {
            @Override
            public Language getLanguage(String code) {
                return null;
            }

            @Override
            public List<Language> loadLanguage() {
                return List.of();
            }
        };
        PdfService pdfService = file -> file;
        TemplateRenderer templateRenderer = new TemplateRenderer(new ExportViewBuilder(pidService, languageService, storage, properties), new DefaultResourceLoader(), properties);
        return new WordFileTransformerService(properties, pdfService, storage, new DefaultResourceLoader(),
                new org.opencdmp.filetransformer.docx.service.wordfiletransformer.validation.DocxTemplateValidationService(), templateRenderer);
    }

    // ---------------------------------------------------------------- fixtures

    private PlanModel plan(PlanStatus status) {
        PlanModel plan = new PlanModel();
        plan.setId(UUID.randomUUID());
        plan.setLabel("Testing Research Plan");
        plan.setVersion((short) 2);
        plan.setDescription("<p>A plan about automated regression testing.</p>");
        plan.setAccessType(PlanAccessType.Public);
        plan.setLanguage("en");

        PlanStatusModel statusModel = new PlanStatusModel();
        statusModel.setId(UUID.randomUUID());
        statusModel.setName(status == PlanStatus.Finalized ? "Finalized" : "Draft");
        statusModel.setInternalStatus(status);
        plan.setStatus(statusModel);

        UserModel creator = new UserModel();
        creator.setId(UUID.randomUUID());
        creator.setName("Ada Researcher");
        plan.setCreator(creator);

        // blueprint: one root section -> one field set -> free-text + rich-text questions
        LabelDataModel textData = new LabelDataModel();
        textData.setFieldType(FieldType.FREE_TEXT);
        textData.setLabel("What data will you collect?");
        FieldModel textField = new FieldModel();
        textField.setId(TEXT_FIELD_ID);
        textField.setOrdinal(0);
        textField.setIncludeInExport(Boolean.TRUE);
        textField.setData(textData);

        LabelDataModel richData = new LabelDataModel();
        richData.setFieldType(FieldType.RICH_TEXT_AREA);
        richData.setLabel("Describe the methodology");
        FieldModel richField = new FieldModel();
        richField.setId(RICH_FIELD_ID);
        richField.setOrdinal(1);
        richField.setIncludeInExport(Boolean.TRUE);
        richField.setData(richData);

        FieldSetModel fieldSet = new FieldSetModel();
        fieldSet.setId(FIELDSET_ID);
        fieldSet.setOrdinal(0);
        fieldSet.setTitle("Data Collection");
        fieldSet.setFields(new ArrayList<>(List.of(textField, richField)));

        // a visible question that is never answered: hidden by default, placeholder on demand
        LabelDataModel emptyData = new LabelDataModel();
        emptyData.setFieldType(FieldType.FREE_TEXT);
        emptyData.setLabel("Backup frequency");
        FieldModel emptyField = new FieldModel();
        emptyField.setId(EMPTY_FIELD_ID);
        emptyField.setOrdinal(0);
        emptyField.setIncludeInExport(Boolean.TRUE);
        emptyField.setData(emptyData);

        FieldSetModel emptyFieldSet = new FieldSetModel();
        emptyFieldSet.setId(EMPTY_FIELDSET_ID);
        emptyFieldSet.setOrdinal(1);
        emptyFieldSet.setTitle("Backup Strategy");
        emptyFieldSet.setFields(new ArrayList<>(List.of(emptyField)));

        RootSectionModel section = new RootSectionModel();
        section.setId(SECTION_ID);
        section.setOrdinal(1);
        section.setTitle("Data Management");
        section.setFieldSets(new ArrayList<>(List.of(fieldSet, emptyFieldSet)));

        DefinitionModel definition = new DefinitionModel();
        definition.setSections(new ArrayList<>(List.of(section)));

        PlanBlueprintModel blueprint = new PlanBlueprintModel();
        blueprint.setId(UUID.randomUUID());
        blueprint.setLabel("Test Blueprint");
        blueprint.setDefinition(definition);
        plan.setPlanBlueprint(blueprint);

        // values
        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel textValue = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        textValue.setTextValue("Sensor measurements and interviews");
        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel richValue = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        richValue.setTextValue("<p>This is <b>rich</b> methodology content.</p>");

        PropertyDefinitionFieldSetItemModel item = new PropertyDefinitionFieldSetItemModel();
        item.setOrdinal(0);
        item.setFields(Map.of(TEXT_FIELD_ID, textValue, RICH_FIELD_ID, richValue));

        PropertyDefinitionFieldSetModel fieldSetValues = new PropertyDefinitionFieldSetModel();
        fieldSetValues.setItems(new ArrayList<>(List.of(item)));

        PropertyDefinitionFieldSetItemModel emptyItem = new PropertyDefinitionFieldSetItemModel();
        emptyItem.setOrdinal(0);
        emptyItem.setFields(Map.of());
        PropertyDefinitionFieldSetModel emptyFieldSetValues = new PropertyDefinitionFieldSetModel();
        emptyFieldSetValues.setItems(new ArrayList<>(List.of(emptyItem)));

        PlanPropertiesModel planProperties = new PlanPropertiesModel();
        planProperties.setFieldSets(new java.util.HashMap<>(Map.of(FIELDSET_ID, fieldSetValues, EMPTY_FIELDSET_ID, emptyFieldSetValues)));
        plan.setProperties(planProperties);

        List<VisibilityStateModel> states = visibility(SECTION_ID, FIELDSET_ID, TEXT_FIELD_ID, RICH_FIELD_ID);
        states.addAll(visibility(SECTION_ID, EMPTY_FIELDSET_ID, EMPTY_FIELD_ID));
        plan.setVisibilityStates(states);
        return plan;
    }

    private List<VisibilityStateModel> visibility(String sectionId, String fieldSetId, String... fieldIds) {
        List<VisibilityStateModel> states = new ArrayList<>();
        VisibilityStateModel section = new VisibilityStateModel();
        section.setFieldId(sectionId);
        section.setOrdinal(null);
        section.setVisible(true);
        states.add(section);
        for (Integer ordinal : new Integer[]{null, 0}) {
            VisibilityStateModel fieldSet = new VisibilityStateModel();
            fieldSet.setFieldId(fieldSetId);
            fieldSet.setOrdinal(ordinal);
            fieldSet.setVisible(true);
            states.add(fieldSet);
            for (String fieldId : fieldIds) {
                VisibilityStateModel field = new VisibilityStateModel();
                field.setFieldId(fieldId);
                field.setOrdinal(ordinal);
                field.setVisible(true);
                states.add(field);
            }
        }
        return states;
    }

    private DescriptionModel description(DescriptionStatus status) {
        DescriptionModel description = new DescriptionModel();
        description.setId(UUID.randomUUID());
        description.setLabel("Survey Dataset");
        description.setDescription("<p>Dataset of survey responses.</p>");
        description.setTags(List.of("survey", "dataset"));
        description.setSectionId(SECTION_ID);

        DescriptionStatusModel statusModel = new DescriptionStatusModel();
        statusModel.setId(UUID.randomUUID());
        statusModel.setName(status == DescriptionStatus.Finalized ? "Finalized" : "Draft");
        statusModel.setInternalStatus(status);
        description.setStatus(statusModel);

        LabelDataModel data = new LabelDataModel();
        data.setFieldType(FieldType.FREE_TEXT);
        data.setLabel("Dataset format");
        FieldModel field = new FieldModel();
        field.setId("df-1");
        field.setOrdinal(0);
        field.setIncludeInExport(Boolean.TRUE);
        field.setData(data);

        FieldSetModel fieldSet = new FieldSetModel();
        fieldSet.setId("dfs-1");
        fieldSet.setOrdinal(0);
        fieldSet.setTitle("Format");
        fieldSet.setFields(new ArrayList<>(List.of(field)));

        SectionModel section = new SectionModel();
        section.setId("dsec-1");
        section.setOrdinal(0);
        section.setTitle("Dataset Details");
        section.setFieldSets(new ArrayList<>(List.of(fieldSet)));

        org.opencdmp.commonmodels.models.descriptiotemplate.DefinitionModel definition = new org.opencdmp.commonmodels.models.descriptiotemplate.DefinitionModel();
        definition.setSections(new ArrayList<>(List.of(section)));

        DescriptionTemplateModel template = new DescriptionTemplateModel();
        template.setId(UUID.randomUUID());
        template.setLabel("Dataset Template");
        template.setDefinition(definition);
        description.setDescriptionTemplate(template);

        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel value = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        value.setTextValue("Comma separated values");
        PropertyDefinitionFieldSetItemModel item = new PropertyDefinitionFieldSetItemModel();
        item.setOrdinal(0);
        item.setFields(Map.of("df-1", value));
        PropertyDefinitionFieldSetModel fieldSetValues = new PropertyDefinitionFieldSetModel();
        fieldSetValues.setItems(new ArrayList<>(List.of(item)));

        PropertyDefinitionModel propertyDefinition = new PropertyDefinitionModel();
        propertyDefinition.setFieldSets(Map.of("dfs-1", fieldSetValues));
        description.setProperties(propertyDefinition);

        description.setVisibilityStates(visibility("dsec-1", "dfs-1", "df-1"));
        description.setPlan(plan(PlanStatus.Draft));
        return description;
    }

    // ---------------------------------------------------------------- helpers

    private XWPFDocument parse(FileEnvelopeModel envelope) throws Exception {
        assertNotNull(envelope.getFile(), "expected inline file bytes");
        return new XWPFDocument(new ByteArrayInputStream(envelope.getFile()));
    }

    private boolean hasWatermark(XWPFDocument document) {
        return hasWatermark(document, "DRAFT");
    }

    private boolean hasWatermark(XWPFDocument document, String text) {
        for (XWPFHeader header : document.getHeaderList()) {
            if (header._getHdrFtr().xmlText().contains("string=\"" + text + "\"")) return true;
        }
        return false;
    }

    /**
     * Word rejects a header whose watermark paragraph got nested inside another {@code <w:p>}
     * instead of being inserted as a sibling (a past regression from mis-using XmlObject#set
     * instead of moving the parsed XML's contents into the target node).
     */
    private void assertNoNestedParagraphs(XWPFDocument document) {
        for (XWPFHeader header : document.getHeaderList()) {
            assertFalse(header._getHdrFtr().xmlText().matches("(?s).*<w:p(?: [^>]*)?><w:p(?: [^>]*)?>.*"),
                    "header contains a paragraph nested directly inside another paragraph");
        }
    }

    // ---------------------------------------------------------------- tests

    @Test
    void exportPlanFinalizedHasNoWatermark() throws Exception {
        FileEnvelopeModel envelope = service().exportPlan(plan(PlanStatus.Finalized), "docx");
        try (XWPFDocument document = parse(envelope)) {
            assertFalse(hasWatermark(document), "finalized plan must not carry a watermark");
        }
    }

    @Test
    void exportDescriptionResolvesContentAndStyles() throws Exception {
        FileEnvelopeModel envelope = service().exportDescription(description(DescriptionStatus.Draft), "docx");
        try (XWPFDocument document = parse(envelope)) {
            assertTrue(document.getStyles().styleExist(TemplateStyles.ANSWER));

            String text = bodyText(document);
            assertTrue(text.contains("Survey Dataset"), "description title missing");
            assertTrue(text.contains("Dataset Details"), "section heading missing");
            assertTrue(text.contains("Comma separated values"), "answer missing");
            assertFalse(text.contains("{{"), "unrendered tag left in document");

            assertTrue(hasWatermark(document), "draft description should carry the DRAFT watermark");
            assertEquals("Survey Dataset", document.getProperties().getCoreProperties().getTitle());
            assertEquals("survey, dataset", document.getProperties().getCoreProperties().getKeywords());
        }
    }

    @Test
    void exportOptionsApplyBrandingAndDocumentSettings() throws Exception {
        PlanModel plan = plan(PlanStatus.Finalized);
        int picturesWithoutLogo;
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            picturesWithoutLogo = document.getAllPictures().size();
        }

        // branding via definition/tenant scope: accent color + cover logo
        FileEnvelopeModel logo = new FileEnvelopeModel();
        logo.setFilename("logo.png");
        logo.setMimeType("image/png");
        logo.setFile(tinyPng());
        PluginFieldModel accentField = new PluginFieldModel();
        accentField.setCode(WordExportOptions.FIELD_ACCENT_COLOR);
        accentField.setTextValue("#aa3366");
        PluginFieldModel logoField = new PluginFieldModel();
        logoField.setCode(WordExportOptions.FIELD_COVER_LOGO);
        logoField.setFile(logo);
        // page size and unanswered questions are organisation settings, set with the branding
        PluginFieldModel pageSizeField = new PluginFieldModel();
        pageSizeField.setCode(WordExportOptions.FIELD_PAGE_SIZE);
        pageSizeField.setTextValue("Letter");
        PluginFieldModel unansweredField = new PluginFieldModel();
        unansweredField.setCode(WordExportOptions.FIELD_UNANSWERED_QUESTIONS);
        unansweredField.setTextValue("placeholder");
        PluginModel plugin = new PluginModel();
        plugin.setCode(properties.getTransformerId());
        plugin.setType(PluginType.FileTransformer);
        plugin.setFields(new ArrayList<>(List.of(accentField, logoField, pageSizeField, unansweredField)));
        plan.getPlanBlueprint().getDefinition().setPlugins(new ArrayList<>(List.of(plugin)));

        // a value a user stored before these became organisation settings no longer overrides them
        PluginUserFieldModel stalePageSize = new PluginUserFieldModel();
        stalePageSize.setCode(WordExportOptions.FIELD_PAGE_SIZE);
        stalePageSize.setTextValue("A4");
        PluginUserModel userPlugin = new PluginUserModel();
        userPlugin.setCode(properties.getTransformerId());
        userPlugin.setType(PluginType.FileTransformer);
        userPlugin.setFields(new ArrayList<>(List.of(stalePageSize)));
        UserAdditionalInfoModel additionalInfo = new UserAdditionalInfoModel();
        additionalInfo.setPluginConfigurations(new ArrayList<>(List.of(userPlugin)));
        plan.getCreator().setAdditionalInfo(additionalInfo);

        FileEnvelopeModel envelope = service().exportPlan(plan, "docx");
        try (XWPFDocument document = parse(envelope)) {
            assertEquals("12240", String.valueOf(document.getDocument().getBody().getSectPr().getPgSz().getW()),
                    "page size should be Letter");
            assertTrue(document.getStyles().getCtStyles().xmlText().contains("AA3366"),
                    "accent color not applied to style definitions");

            String text = bodyText(document);
            assertTrue(text.contains("Backup Strategy"), "unanswered question heading should be kept in placeholder mode");
            assertTrue(text.contains("Not answered"), "placeholder text missing");
            assertFalse(text.contains("{{"), "unrendered tag left in document");

            assertTrue(document.getAllPictures().size() > picturesWithoutLogo, "cover logo not embedded");
        }
    }

    @Test
    void watermarkShowsTheStatusUnlessATextIsSet() throws Exception {
        PlanModel plan = planWithAsset();
        plan.getStatus().setName("Under Review");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertTrue(hasWatermark(document, "UNDER REVIEW"), "a plan that is not finalized should carry its status");
        }
        plan.getStatus().setName(null);
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertTrue(hasWatermark(document, "DRAFT"), "without a status name the DRAFT label applies");
        }
        plan.getStatus().setName("Under Review");
        layout(plan, WordExportOptions.FIELD_DRAFT_WATERMARK_TEXT, "Confidential");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertTrue(hasWatermark(document, "Confidential"), "a configured text replaces the status");
            assertFalse(hasWatermark(document, "UNDER REVIEW"), "the status should not appear when a text is set");
        }
    }

    @Test
    void selectiveExportIncludesOnlyChosenSections() throws Exception {
        PlanModel plan = plan(PlanStatus.Finalized);

        // second root section with an answered question
        LabelDataModel archiveData = new LabelDataModel();
        archiveData.setFieldType(FieldType.FREE_TEXT);
        archiveData.setLabel("Where is the data archived?");
        FieldModel archiveField = new FieldModel();
        archiveField.setId("f-arch");
        archiveField.setOrdinal(0);
        archiveField.setIncludeInExport(Boolean.TRUE);
        archiveField.setData(archiveData);
        FieldSetModel archiveFieldSet = new FieldSetModel();
        archiveFieldSet.setId("fs-3");
        archiveFieldSet.setOrdinal(0);
        archiveFieldSet.setTitle("Archiving Details");
        archiveFieldSet.setFields(new ArrayList<>(List.of(archiveField)));
        RootSectionModel archiveSection = new RootSectionModel();
        archiveSection.setId("sec-2");
        archiveSection.setOrdinal(2);
        archiveSection.setTitle("Archiving");
        archiveSection.setFieldSets(new ArrayList<>(List.of(archiveFieldSet)));
        plan.getPlanBlueprint().getDefinition().getSections().add(archiveSection);

        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel archiveValue = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        archiveValue.setTextValue("In the institutional repository");
        PropertyDefinitionFieldSetItemModel archiveItem = new PropertyDefinitionFieldSetItemModel();
        archiveItem.setOrdinal(0);
        archiveItem.setFields(Map.of("f-arch", archiveValue));
        PropertyDefinitionFieldSetModel archiveValues = new PropertyDefinitionFieldSetModel();
        archiveValues.setItems(new ArrayList<>(List.of(archiveItem)));
        plan.getProperties().getFieldSets().put("fs-3", archiveValues);
        plan.getVisibilityStates().addAll(visibility("sec-2", "fs-3", "f-arch"));

        // without selection both sections render
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Data Collection"), "first section missing in full export");
            assertTrue(text.contains("Archiving Details"), "second section missing in full export");
        }

        // user selects only the first section
        PluginUserFieldModel includeSections = new PluginUserFieldModel();
        includeSections.setCode(WordExportOptions.FIELD_INCLUDE_SECTIONS);
        includeSections.setTextValue(SECTION_ID);
        PluginUserModel userPlugin = new PluginUserModel();
        userPlugin.setCode(properties.getTransformerId());
        userPlugin.setType(PluginType.FileTransformer);
        userPlugin.setFields(new ArrayList<>(List.of(includeSections)));
        UserAdditionalInfoModel additionalInfo = new UserAdditionalInfoModel();
        additionalInfo.setPluginConfigurations(new ArrayList<>(List.of(userPlugin)));
        plan.getCreator().setAdditionalInfo(additionalInfo);

        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Data Collection"), "included section missing");
            assertTrue(text.contains("Sensor measurements and interviews"), "included answer missing");
            assertFalse(text.contains("Archiving Details"), "excluded section rendered");
            assertFalse(text.contains("In the institutional repository"), "excluded answer rendered");
        }
    }

    /** A template made for the original engine, which is no longer supported. */
    private byte[] oldPlaceholderTemplate() throws Exception {
        try (XWPFDocument custom = new XWPFDocument(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            custom.createParagraph().createRun().setText("'{OPENCDMP.PLAN.TITLE}'");
            custom.createParagraph().createRun().setText("'{OPENCDMP.PLAN.FORM}'");
            custom.write(out);
            return out.toByteArray();
        }
    }

    @Test
    void validateTemplateRejectsOldPlaceholders() throws Exception {
        FileEnvelopeModel envelope = new FileEnvelopeModel();
        envelope.setFilename("old.docx");
        envelope.setFile(oldPlaceholderTemplate());

        var result = service().validateTemplate(envelope, PluginEntityType.Plan);
        assertNotNull(result);
        assertFalse(result.isValid(), "a template with '{OPENCDMP.X}' placeholders must be invalid");
        assertTrue(result.getIssues().stream().anyMatch(x -> "unsupported-placeholders".equals(x.getCode())), describe(result));
    }

    private String describe(org.opencdmp.filetransformerbase.models.misc.TemplateValidationResultModel result) {
        return result.getIssues().stream().map(x -> x.getCode() + ": " + x.getMessage()).reduce("", (a, b) -> a + " | " + b);
    }

    private byte[] tinyPng() throws Exception {
        java.awt.image.BufferedImage image = new java.awt.image.BufferedImage(8, 8, java.awt.image.BufferedImage.TYPE_INT_RGB);
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            javax.imageio.ImageIO.write(image, "png", out);
            return out.toByteArray();
        }
    }

    @Test
    void customTemplateWithOldPlaceholdersFallsBackToTheBuiltInTemplate() throws Exception {
        PlanModel plan = plan(PlanStatus.Draft);
        FileEnvelopeModel templateEnvelope = new FileEnvelopeModel();
        templateEnvelope.setFilename("old.docx");
        templateEnvelope.setFile(oldPlaceholderTemplate());
        PluginFieldModel pluginField = new PluginFieldModel();
        pluginField.setCode("plan-docx-template-id");
        pluginField.setFile(templateEnvelope);
        PluginModel plugin = new PluginModel();
        plugin.setCode(properties.getTransformerId());
        plugin.setType(PluginType.FileTransformer);
        plugin.setFields(new ArrayList<>(List.of(pluginField)));
        plan.getPlanBlueprint().getDefinition().setPlugins(new ArrayList<>(List.of(plugin)));

        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertFalse(text.contains("{OPENCDMP"), "old placeholders must not reach the export: " + text);
            assertTrue(text.contains("Testing Research Plan"), "built-in template not used");
            assertTrue(text.contains("Sensor measurements and interviews"), "plan content missing");
            for (String styleId : TemplateStyles.CONTRACT) {
                assertTrue(document.getStyles().styleExist(styleId), "missing contract style: " + styleId);
            }
        }
    }

    // ---------------------------------------------------------------- '{{...}}' templates

    private PluginModel layout(PlanModel plan, String... codeValues) {
        List<PluginFieldModel> fields = new ArrayList<>();
        for (int i = 0; i < codeValues.length; i += 2) {
            PluginFieldModel field = new PluginFieldModel();
            field.setCode(codeValues[i]);
            field.setTextValue(codeValues[i + 1]);
            fields.add(field);
        }
        PluginModel plugin = new PluginModel();
        plugin.setCode(properties.getTransformerId());
        plugin.setType(PluginType.FileTransformer);
        plugin.setFields(fields);
        plan.getPlanBlueprint().getDefinition().setPlugins(new ArrayList<>(List.of(plugin)));
        return plugin;
    }

    /** A plan with one asset in its only section. */
    private PlanModel planWithAsset() {
        PlanModel plan = plan(PlanStatus.Draft);
        DescriptionModel asset = description(DescriptionStatus.Draft);
        asset.setCreatedAt(java.time.Instant.parse("2026-01-01T00:00:00Z"));
        asset.setPlan(null);
        plan.setDescriptions(new ArrayList<>(List.of(asset)));
        return plan;
    }

    private String bodyText(XWPFDocument document) {
        StringBuilder text = new StringBuilder();
        for (var paragraph : document.getParagraphs()) text.append(paragraph.getText()).append('\n');
        for (var table : document.getTables()) text.append(table.getText()).append('\n');
        return text.toString();
    }

    @Test
    void templatePlanExportRendersStructureQuestionsAndStyles() throws Exception {
        try (XWPFDocument document = parse(service().exportPlan(planWithAsset(), "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Testing Research Plan"), "cover title missing");
            assertTrue(text.contains("1\u2002Data Management"), "numbered section heading missing: " + text);
            assertTrue(text.contains("What data will you collect?"), "question label missing");
            assertTrue(text.contains("Describe the methodology"), "question label missing");
            assertTrue(text.contains("Sensor measurements and interviews"), "free-text answer missing");
            assertTrue(text.contains("rich"), "rich-text answer missing");
            assertFalse(text.contains("Backup Strategy"), "unanswered question should be hidden by default");
            assertTrue(text.contains("Survey Dataset"), "asset title missing");
            assertTrue(text.contains("Comma separated values"), "asset answer missing");
            assertFalse(text.contains("{{"), "unrendered tag left: " + text);

            for (String styleId : TemplateStyles.CONTRACT) {
                assertTrue(document.getStyles().styleExist(styleId), "missing contract style: " + styleId);
            }
            assertTrue(document.getParagraphs().stream().anyMatch(p -> TemplateStyles.QUESTION.equals(p.getStyle())), "no question label paragraph");
            assertTrue(document.getParagraphs().stream().anyMatch(p -> TemplateStyles.ANSWER.equals(p.getStyle())), "no answer paragraph");
            assertTrue(document.getDocument().xmlText().contains(" TOC "), "table of contents field missing");
            assertFalse(document.getSettings().getCTSettings().isSetUpdateFields(), "settings.xml must not force field updates on open");
            assertTrue(hasWatermark(document), "draft plan should carry the DRAFT watermark");
            assertNoNestedParagraphs(document);
            assertEquals("Testing Research Plan", document.getProperties().getCoreProperties().getTitle());
        }
    }

    @Test
    void templateExportKeepsAssetsOfASectionWithoutQuestions() throws Exception {
        // a blueprint section that only holds assets: the platform reports it as not visible, as it has no visible questions
        PlanModel plan = planWithAsset();
        RootSectionModel assetsSection = new RootSectionModel();
        assetsSection.setId("sec-assets");
        assetsSection.setOrdinal(2);
        assetsSection.setTitle("Research outputs");
        assetsSection.setFieldSets(new ArrayList<>());
        plan.getPlanBlueprint().getDefinition().getSections().add(assetsSection);
        VisibilityStateModel hidden = new VisibilityStateModel();
        hidden.setFieldId("sec-assets");
        hidden.setOrdinal(null);
        hidden.setVisible(false);
        plan.getVisibilityStates().add(hidden);
        plan.getDescriptions().getFirst().setSectionId("sec-assets");

        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Survey Dataset"), "asset of a section without questions missing: " + text);
            assertTrue(text.contains("Comma separated values"), "that asset's answers missing");
        }
    }

    @Test
    void templateAssetPlacementOrdersAssetsAroundSections() throws Exception {
        PlanModel plan = planWithAsset();
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.indexOf("Comma separated values") > text.indexOf("Sensor measurements and interviews"), "asset should follow its section by default");
        }
        layout(plan, WordExportOptions.FIELD_LAYOUT_ASSET_PLACEMENT, "before");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Assets"), "asset group heading missing");
            assertTrue(text.indexOf("Comma separated values") < text.indexOf("Sensor measurements and interviews"), "assets should come first: " + text);
        }
    }

    @Test
    void templateAssetsContinueTheNumberingOfWhereTheyAppear() throws Exception {
        PlanModel plan = planWithAsset();
        // after their section: the asset is a sub-chapter of section 1, and its own sections continue from there
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("1 Data Management"), "section number missing: " + text);
            // the section's own question group is 1.1, so its asset follows as 1.2
            assertTrue(text.contains("1.2 Survey Dataset"), "asset should be numbered after its section's questions: " + text);
            assertTrue(text.contains("1.2.1 "), "the asset's contents should continue its number: " + text);
        }
        // grouped after the sections: the Assets chapter takes the next number
        layout(plan, WordExportOptions.FIELD_LAYOUT_ASSET_PLACEMENT, "after");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("1 Data Management"), "section number missing: " + text);
            assertTrue(text.contains("2 Assets"), "the Assets chapter should be numbered after the sections: " + text);
            assertTrue(text.contains("2.1 Survey Dataset"), "asset should be numbered under the Assets chapter: " + text);
        }
        // grouped before the sections: the Assets chapter is 1 and the sections follow
        layout(plan, WordExportOptions.FIELD_LAYOUT_ASSET_PLACEMENT, "before");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("1 Assets"), "the Assets chapter should come first: " + text);
            assertTrue(text.contains("1.1 Survey Dataset"), "asset should be numbered under the Assets chapter: " + text);
            assertTrue(text.contains("2 Data Management"), "sections should be numbered after the Assets chapter: " + text);
        }
    }

    @Test
    void templateAccentColourAlsoRecoloursTheCoverStripe() throws Exception {
        PlanModel plan = planWithAsset();
        layout(plan, WordExportOptions.FIELD_ACCENT_COLOR, "#aa3366");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String xml = document.getDocument().xmlText();
            assertTrue(xml.contains("OpencdmpCoverStripe"), "cover stripe missing");
            assertTrue(xml.contains("fillcolor=\"#AA3366\""), "cover stripe should take the accent colour");
            assertFalse(xml.contains("fillcolor=\"#2457C5\""), "the template accent should be replaced everywhere");
        }
    }

    @Test
    void templateLayoutOptionsChangeTheDocument() throws Exception {
        PlanModel plan = planWithAsset();
        layout(plan,
                WordExportOptions.FIELD_LAYOUT_COVER, "false",
                WordExportOptions.FIELD_LAYOUT_TOC, "false",
                WordExportOptions.FIELD_LAYOUT_NUMBERING, "false",
                WordExportOptions.FIELD_LAYOUT_QUESTION_STYLE, "table",
                WordExportOptions.FIELD_ACCENT_COLOR, "#aa3366");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertFalse(document.getParagraphs().stream().anyMatch(p -> "Title".equals(p.getStyle())), "cover page should be left out");
            assertFalse(document.getDocument().xmlText().contains(" TOC "), "table of contents should be left out");
            assertFalse(text.contains("1\u2002"), "numbering should be off");
            assertTrue(document.getTables().stream().anyMatch(t -> t.getText().contains("What data will you collect?") && t.getText().contains("Sensor measurements")),
                    "questions should render as a question / answer table");
            String styles = document.getStyles().getCtStyles().xmlText();
            assertTrue(styles.contains("AA3366"), "accent color not applied");
            assertFalse(styles.contains(TemplateStyles.ACCENT_COLOR), "template accent color left in styles");
        }
    }

    @Test
    void templateResolvesOptionLabelsInternalEntriesAndUnansweredPlaceholders() throws Exception {
        PlanModel plan = planWithAsset();
        DescriptionModel asset = plan.getDescriptions().getFirst();
        FieldSetModel fieldSet = asset.getDescriptionTemplate().getDefinition().getSections().getFirst().getFieldSets().getFirst();

        org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SelectDataModel select = new org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SelectDataModel();
        select.setFieldType(FieldType.SELECT);
        select.setLabel("Access level");
        select.setMultipleSelect(false);
        org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SelectDataModel.OptionModel option = new org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SelectDataModel.OptionModel();
        option.setLabel("Open access");
        option.setValue("open");
        select.setOptions(List.of(option));
        FieldModel selectField = new FieldModel();
        selectField.setId("df-select");
        selectField.setOrdinal(1);
        selectField.setIncludeInExport(true);
        selectField.setData(select);

        org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.LabelAndMultiplicityDataModel internal = new org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.LabelAndMultiplicityDataModel();
        internal.setFieldType(FieldType.INTERNAL_ENTRIES_DESCRIPTIONS);
        internal.setLabel("Related assets");
        FieldModel internalField = new FieldModel();
        internalField.setId("df-internal");
        internalField.setOrdinal(2);
        internalField.setIncludeInExport(true);
        internalField.setData(internal);

        LabelDataModel unanswered = new LabelDataModel();
        unanswered.setFieldType(FieldType.FREE_TEXT);
        unanswered.setLabel("Retention period");
        FieldModel unansweredField = new FieldModel();
        unansweredField.setId("df-empty");
        unansweredField.setOrdinal(3);
        unansweredField.setIncludeInExport(true);
        unansweredField.setData(unanswered);
        fieldSet.getFields().addAll(List.of(selectField, internalField, unansweredField));

        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel selectValue = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        selectValue.setTextListValue(List.of("open"));
        org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel internalValue = new org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel();
        internalValue.setTextListValue(List.of(asset.getId().toString()));
        PropertyDefinitionFieldSetItemModel item = asset.getProperties().getFieldSets().get("dfs-1").getItems().getFirst();
        java.util.Map<String, org.opencdmp.commonmodels.models.plandescriptionfields.FieldModel> values = new java.util.HashMap<>(item.getFields());
        values.put("df-select", selectValue);
        values.put("df-internal", internalValue);
        item.setFields(values);
        List<VisibilityStateModel> states = new ArrayList<>(asset.getVisibilityStates());
        states.addAll(visibility("dsec-1", "dfs-1", "df-select", "df-internal", "df-empty"));
        asset.setVisibilityStates(states);

        // an organisation setting: unanswered questions are kept with a placeholder
        layout(plan, WordExportOptions.FIELD_UNANSWERED_QUESTIONS, "placeholder");

        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Open access"), "select value should be shown by its option label");
            assertFalse(text.contains(asset.getId().toString()), "internal entry id should be replaced by the asset title");
            assertTrue(text.contains("Related assets\nSurvey Dataset"), "internal entry should resolve to the asset title: " + text);
            assertTrue(text.contains("Retention period") && text.contains("Not answered"), "unanswered question should get a placeholder");
        }
    }

    @Test
    void templateAssetExportUsesThePlanLanguage() throws Exception {
        DescriptionModel asset = description(DescriptionStatus.Finalized);
        asset.getPlan().setLanguage("el");
        try (XWPFDocument document = parse(service().exportDescription(asset, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("Survey Dataset"), "asset title missing");
            assertTrue(text.contains("Comma separated values"), "asset answer missing");
            assertTrue(text.contains("Μέρος του Testing Research Plan"), "greek labels expected: " + text);
            assertFalse(hasWatermark(document), "finalized asset must not carry a watermark");
        }
    }

    @Test
    void customTemplateGetsMissingStylesAndGlobalNamesInsideLoops() throws Exception {
        byte[] customTemplate;
        try (XWPFDocument custom = new XWPFDocument(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            custom.createParagraph().createRun().setText("{{?chapters}}");
            custom.createParagraph().createRun().setText("{{?section}}");
            custom.createParagraph().createRun().setText("{{heading}} of {{plan.title}}");
            custom.createParagraph().createRun().setText("{{>body}}");
            custom.createParagraph().createRun().setText("{{/section}}");
            custom.createParagraph().createRun().setText("{{/chapters}}");
            custom.write(out);
            customTemplate = out.toByteArray();
        }
        PlanModel plan = plan(PlanStatus.Draft);
        FileEnvelopeModel templateEnvelope = new FileEnvelopeModel();
        templateEnvelope.setFilename("custom.docx");
        templateEnvelope.setFile(customTemplate);
        PluginModel plugin = layout(plan);
        PluginFieldModel templateField = new PluginFieldModel();
        templateField.setCode("plan-docx-template-id");
        templateField.setFile(templateEnvelope);
        plugin.getFields().add(templateField);

        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(text.contains("1\u2002Data Management of Testing Research Plan"), "global name not resolved inside a loop: " + text);
            assertTrue(text.contains("Sensor measurements and interviews"), "body not rendered");
            for (String styleId : List.of(TemplateStyles.QUESTION, TemplateStyles.ANSWER, TemplateStyles.TABLE_LINED)) {
                assertTrue(document.getStyles().styleExist(styleId), "style not copied into custom template: " + styleId);
            }
            assertTrue(hasWatermark(document), "watermark should be added even when the template has no headers");
            assertNoNestedParagraphs(document);
        }
    }

    @Test
    void validateTemplateAcceptsDefaultTemplatesAndFlagsTemplateMistakes() throws Exception {
        for (String[] builtIn : new String[][]{{"classpath:documents/plan.docx", "Plan"}, {"classpath:documents/description.docx", "Description"}}) {
            FileEnvelopeModel envelope = new FileEnvelopeModel();
            try (var in = new DefaultResourceLoader().getResource(builtIn[0]).getInputStream()) {
                envelope.setFile(in.readAllBytes());
            }
            var result = service().validateTemplate(envelope, PluginEntityType.valueOf(builtIn[1]));
            assertTrue(result.isValid() && result.getIssues().isEmpty(), builtIn[0] + " should have no issues: " + describe(result));
        }

        byte[] template;
        try (XWPFDocument custom = new XWPFDocument(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            custom.createParagraph().createRun().setText("{{plan.title}} {{planTitle}}");
            custom.write(out);
            template = out.toByteArray();
        }
        FileEnvelopeModel envelope = new FileEnvelopeModel();
        envelope.setFile(template);
        var result = service().validateTemplate(envelope, PluginEntityType.Plan);
        assertTrue(result.isValid(), "warnings only: " + describe(result));
        assertTrue(result.getIssues().stream().anyMatch(x -> "unknown-tag".equals(x.getCode())), describe(result));
        assertTrue(result.getIssues().stream().anyMatch(x -> "missing-body".equals(x.getCode())), describe(result));
    }

    @Test
    void coverFactsFollowTheConfiguredRows() throws Exception {
        PlanModel plan = planWithAsset();
        org.opencdmp.commonmodels.models.referencetype.ReferenceTypeModel grants = new org.opencdmp.commonmodels.models.referencetype.ReferenceTypeModel();
        grants.setCode("grants");
        grants.setName("Grant");
        org.opencdmp.commonmodels.models.reference.ReferenceModel grant = new org.opencdmp.commonmodels.models.reference.ReferenceModel();
        grant.setId(UUID.randomUUID());
        grant.setLabel("Open Science Grant");
        grant.setType(grants);
        org.opencdmp.commonmodels.models.planreference.PlanReferenceModel planReference = new org.opencdmp.commonmodels.models.planreference.PlanReferenceModel();
        planReference.setReference(grant);
        plan.setReferences(new ArrayList<>(List.of(planReference)));
        org.opencdmp.commonmodels.models.plan.PlanContactModel contact = new org.opencdmp.commonmodels.models.plan.PlanContactModel();
        contact.setFirstName("Ada");
        contact.setLastName("Researcher");
        plan.getProperties().setContacts(List.of(contact));

        // default rows: grant and contact, labelled from the reference type and the contact label
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String facts = document.getTables().getFirst().getText();
            assertTrue(facts.contains("Grant") && facts.contains("Open Science Grant"), "default grant fact missing: " + facts);
            assertTrue(facts.contains("Ada Researcher"), "default contact fact missing: " + facts);
        }

        // configured rows, in order, with a label override and a built-in key; unknown / empty rows are left out
        layout(plan, WordExportOptions.FIELD_LAYOUT_PLAN_FACTS, "version, grants=Project, doi, blueprint");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String facts = document.getTables().getFirst().getText();
            assertTrue(facts.indexOf("Version") < facts.indexOf("Project"), "rows not in configured order: " + facts);
            assertTrue(facts.contains("Open Science Grant") && facts.contains("Test Blueprint"), facts);
            assertFalse(facts.contains("Ada Researcher"), "contact was not configured: " + facts);
            assertFalse(facts.contains("DOI"), "empty rows must be left out: " + facts);
        }
    }

    @Test
    void assetCoverFactsFollowTheConfiguredRows() throws Exception {
        DescriptionModel asset = description(DescriptionStatus.Finalized);
        asset.getDescriptionTemplate().getDefinition().getAllField().getFirst().setSemantics(List.of("rda.dataset.distribution.format"));
        PluginFieldModel field = new PluginFieldModel();
        field.setCode(WordExportOptions.FIELD_LAYOUT_ASSET_FACTS);
        field.setTextValue("sem:rda.dataset.distribution.format, template=Form");
        PluginModel plugin = new PluginModel();
        plugin.setCode(properties.getTransformerId());
        plugin.setType(PluginType.FileTransformer);
        plugin.setFields(new ArrayList<>(List.of(field)));
        asset.getDescriptionTemplate().getDefinition().setPlugins(new ArrayList<>(List.of(plugin)));

        try (XWPFDocument document = parse(service().exportDescription(asset, "docx"))) {
            String facts = document.getTables().getFirst().getText();
            assertTrue(facts.contains("Dataset format") && facts.contains("Comma separated values"), "semantic fact missing: " + facts);
            assertTrue(facts.contains("Form") && facts.contains("Dataset Template"), "relabelled template fact missing: " + facts);
            assertFalse(facts.contains("Status"), "status was not configured: " + facts);
        }
    }

    @Test
    void guidanceCommentsWordingWatermarkAndPageBreaksAreConfigurable() throws Exception {
        PlanModel plan = planWithAsset();
        FieldSetModel fieldSet = plan.getPlanBlueprint().getDefinition().getSections().getFirst().getFieldSets().getFirst();
        fieldSet.setDescription("<p>Describe the data you will collect.</p>");
        plan.getProperties().getFieldSets().get(FIELDSET_ID).setComment("Internal note for the data steward");
        RootSectionModel second = new RootSectionModel();
        second.setId("sec-2");
        second.setOrdinal(2);
        second.setTitle("Second Section");
        second.setFieldSets(new ArrayList<>(List.of()));
        plan.getPlanBlueprint().getDefinition().getSections().add(second);

        // defaults: no guidance, comments included, translated watermark, no section page breaks
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertFalse(text.contains("Describe the data you will collect."), "guidance is off by default");
            assertTrue(text.contains("Internal note for the data steward"), "comments are on by default");
            assertTrue(text.contains("Generated with OpenCDMP"), "generator note expected by default");
            assertTrue(hasWatermark(document));
        }

        layout(plan,
                WordExportOptions.FIELD_LAYOUT_GUIDANCE, "true",
                WordExportOptions.FIELD_LAYOUT_COMMENTS, "false",
                WordExportOptions.FIELD_LAYOUT_LABELS, "planKind=Output Management Plan; generatedWith=",
                WordExportOptions.FIELD_DRAFT_WATERMARK, "false",
                WordExportOptions.FIELD_LAYOUT_SECTION_PAGE_BREAK, "true",
                WordExportOptions.FIELD_LAYOUT_TOC, "false");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            String text = bodyText(document);
            assertTrue(document.getParagraphs().stream().anyMatch(p -> TemplateStyles.GUIDANCE.equals(p.getStyle()) && p.getText().contains("Describe the data")), "guidance missing");
            assertFalse(text.contains("Internal note for the data steward"), "comments should be left out");
            assertTrue(text.contains("Output Management Plan"), "wording override not applied");
            assertFalse(text.contains("Generated with OpenCDMP"), "blanked text still shown");
            assertFalse(text.contains("  ·  Exported"), "blanked generator left a separator behind");
            assertFalse(hasWatermark(document), "watermark should be off");
            // one page break after the cover; the first chapter after it must not add another
            long breaks = document.getDocument().xmlText().split("w:type=\"page\"", -1).length - 1;
            assertEquals(1, breaks, "only the cover page break expected, the first chapter never breaks");
        }

        // a second exported section starts on a new page
        second.getFieldSets().add(fieldSet);
        plan.getVisibilityStates().addAll(visibility("sec-2", FIELDSET_ID, TEXT_FIELD_ID, RICH_FIELD_ID));
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            long breaks = document.getDocument().xmlText().split("w:type=\"page\"", -1).length - 1;
            assertEquals(2, breaks, "the second section should start on a new page");
        }

        // custom watermark text
        layout(plan, WordExportOptions.FIELD_DRAFT_WATERMARK_TEXT, "FOR REVIEW");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertTrue(document.getHeaderList().stream().anyMatch(h -> h._getHdrFtr().xmlText().contains("FOR REVIEW")), "custom watermark text missing");
        }
    }

    @Test
    void declaredDefaultsApplyAndBlankValuesFallThrough() throws Exception {
        ConfigurationField questionStyle = new ConfigurationField(WordExportOptions.FIELD_LAYOUT_QUESTION_STYLE, DataType.Select, "Question Layout", List.of(PluginEntityType.Plan), false);
        questionStyle.setDefaultValue("table");
        List<ConfigurationField> fields = new ArrayList<>(properties.getConfigurationFields());
        fields.add(questionStyle);
        properties.setConfigurationFields(fields);

        // the blueprint stores the field without a value: the declared default applies
        PlanModel plan = planWithAsset();
        layout(plan, WordExportOptions.FIELD_LAYOUT_QUESTION_STYLE, "");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertTrue(document.getTables().stream().anyMatch(t -> t.getText().contains("What data will you collect?")), "declared default (question table) not applied");
        }

        // a configured value still wins over the declared default
        layout(plan, WordExportOptions.FIELD_LAYOUT_QUESTION_STYLE, "stacked");
        try (XWPFDocument document = parse(service().exportPlan(plan, "docx"))) {
            assertFalse(document.getTables().stream().anyMatch(t -> t.getText().contains("What data will you collect?")), "configured value should override the default");
        }
    }
}
