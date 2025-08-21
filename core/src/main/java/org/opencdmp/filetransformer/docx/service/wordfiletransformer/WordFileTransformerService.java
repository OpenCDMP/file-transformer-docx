package org.opencdmp.filetransformer.docx.service.wordfiletransformer;

import gr.cite.tools.exception.MyApplicationException;
import gr.cite.tools.logging.LoggerService;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.opencdmp.commonmodels.enums.*;
import org.opencdmp.commonmodels.models.ConfigurationField;
import org.opencdmp.commonmodels.models.plan.PlanBlueprintValueModel;
import org.opencdmp.commonmodels.models.plan.PlanContactModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.descriptiotemplate.DescriptionTemplateModel;
import org.opencdmp.commonmodels.models.planblueprint.*;
import org.opencdmp.commonmodels.models.planreference.PlanReferenceModel;
import org.opencdmp.commonmodels.models.plugin.PluginFieldModel;
import org.opencdmp.commonmodels.models.plugin.PluginModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerClient;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerConfiguration;
import org.opencdmp.filetransformer.docx.model.enums.FileFormats;
import org.opencdmp.filetransformerbase.models.misc.*;
import org.opencdmp.filetransformer.docx.service.pdf.PdfService;
import org.opencdmp.filetransformer.docx.model.enums.ParagraphStyle;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.visibility.VisibilityServiceImpl;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.word.WordBuilder;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Component;
import org.springframework.web.context.annotation.RequestScope;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import javax.management.InvalidApplicationException;
import java.io.*;
import java.math.BigInteger;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

import static org.apache.poi.xwpf.usermodel.Document.*;
import static org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_WMF;

@Component
@RequestScope
public class WordFileTransformerService implements FileTransformerClient {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(WordFileTransformerService.class));

    private final static List<FileFormat> FILE_FORMATS = List.of(
            new FileFormat(FileFormats.PDF.getValue(), true, "fa-file-pdf-o"),
            new FileFormat(FileFormats.DOCX.getValue(), true, "fa-file-word-o"));

    private final static List<PluginEntityType> FILE_TRANSFORMER_ENTITY_TYPES = List.of(
            PluginEntityType.Plan, PluginEntityType.Description);

    private static final Map<String, Integer> IMAGE_TYPE_MAP = Map.of(
            "image/jpeg", PICTURE_TYPE_JPEG,
            "image/png", PICTURE_TYPE_PNG,
            "image/gif", PICTURE_TYPE_GIF,
            "image/tiff", PICTURE_TYPE_TIFF,
            "image/bmp", PICTURE_TYPE_BMP,
            "image/wmf", PICTURE_TYPE_WMF
    );
    private Integer imageCount;
    private final FileStorageService fileStorageService;
    private final WordFileTransformerServiceProperties wordFileTransformerServiceProperties;
    private final PdfService pdfService;
    private final WordBuilder wordBuilder;
    private final FileStorageService storageService;
    private final MessageSource messageSource;
    private final ResourceLoader resourceLoader;
    @Autowired
    public WordFileTransformerService(
            FileStorageService fileStorageService, WordFileTransformerServiceProperties wordFileTransformerServiceProperties,
            PdfService pdfService, WordBuilder wordBuilder, FileStorageService storageService, MessageSource messageSource, ResourceLoader resourceLoader) {
        this.fileStorageService = fileStorageService;
        this.wordFileTransformerServiceProperties = wordFileTransformerServiceProperties;
	    this.pdfService = pdfService;
	    this.wordBuilder = wordBuilder;
	    this.storageService = storageService;
	    this.messageSource = messageSource;
        this.resourceLoader = resourceLoader;
        this.imageCount = 0;
    }

    @Override
    public FileEnvelopeModel exportPlan(PlanModel plan, String variant) throws IOException, InvalidApplicationException {
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
        return wordFile;
    }

    @Override
    public FileEnvelopeModel exportDescription(DescriptionModel descriptionModel, String variant) throws InvalidApplicationException, IOException {
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
    

    private byte[] buildPlanWordDocument(PlanModel planEntity) throws IOException, InvalidApplicationException {
        if (planEntity == null) throw new MyApplicationException("planEntity required");
        PlanBlueprintModel planBlueprintModel = planEntity.getPlanBlueprint();
        if (planBlueprintModel == null) throw new MyApplicationException("PlanBlueprint required");
        if (planBlueprintModel.getDefinition() == null) throw new MyApplicationException("PlanBlueprint Definition required");
        if (planBlueprintModel.getDefinition().getSections() == null) throw new MyApplicationException("PlanBlueprint Section required");


        XWPFDocument document = null;
        if (planBlueprintModel.getDefinition().getPlugins() != null && !planBlueprintModel.getDefinition().getPlugins().isEmpty()) {
            document = this.getCustomDocument(planBlueprintModel.getDefinition().getPlugins(), PluginEntityType.Plan);
        }
        if (document == null) {
            try {
                Resource resource = resourceLoader.getResource(this.wordFileTransformerServiceProperties.getWordPlanTemplate());
                try(InputStream inputStream = resource.getInputStream()) {
                    document = new XWPFDocument(inputStream);
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new RuntimeException(e);
            }
        }

        this.wordBuilder.fillFirstPage(planEntity, null, document, false);

        int powered_pos = this.wordBuilder.findPosOfPoweredBy(document);
        XWPFParagraph powered_par = null;
        XWPFParagraph argos_img_par = null;
        if (powered_pos != -1) {
            powered_par = document.getParagraphArray(powered_pos);
            argos_img_par = document.getParagraphArray(powered_pos + 1);
        }

        for (SectionModel sectionModel : planBlueprintModel.getDefinition().getSections()) {
            buildPlanSection(planEntity, sectionModel, document);
        }

        if (powered_pos != -1) {
            document.getLastParagraph().setPageBreak(false);
            document.createParagraph();
            document.setParagraph(powered_par, document.getParagraphs().size() - 1);

            document.createParagraph();
            document.setParagraph(argos_img_par, document.getParagraphs().size() - 1);

            document.removeBodyElement(powered_pos + 1);
            document.removeBodyElement(powered_pos + 1);
        }

        this.wordBuilder.fillFooter(planEntity, null, document);
        this.wordBuilder.fillHeader(planEntity, null, document);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        document.write(out);
        byte[] bytes = out.toByteArray();
        out.close();

        return bytes;
    }

    private void buildPlanSection(PlanModel planEntity, SectionModel sectionModel, XWPFDocument document) throws InvalidApplicationException {
        this.wordBuilder.addParagraphContent(sectionModel.getOrdinal() + ". " + sectionModel.getLabel(), document, ParagraphStyle.HEADER1, BigInteger.ZERO, 0);

        if (sectionModel.getFields() != null) {
            sectionModel.getFields().sort(Comparator.comparingInt(FieldModel::getOrdinal));
            for (FieldModel fieldModel : sectionModel.getFields()) {
                buildPlanSectionField(planEntity, document, fieldModel);
            }
        }
        
        final boolean isFinalized = planEntity.getStatus() != null && planEntity.getStatus().getInternalStatus() != null && planEntity.getStatus().getInternalStatus().equals(PlanStatus.Finalized);
        final boolean isPublic = planEntity.getPublicAfter() != null && planEntity.getPublicAfter().isAfter(Instant.now());
        
        List<DescriptionModel> descriptions = planEntity.getDescriptions() == null ? new ArrayList<>() : planEntity.getDescriptions().stream()
                .filter(item -> item.getStatus() != null && (item.getStatus().getInternalStatus() == null || (item.getStatus().getInternalStatus() != null && item.getStatus().getInternalStatus() != DescriptionStatus.Canceled )))
                .filter(item -> !isPublic && !isFinalized || (item.getStatus().getInternalStatus() != null && item.getStatus().getInternalStatus() == DescriptionStatus.Finalized))
                .filter(item -> item.getSectionId().equals(sectionModel.getId()))
                .sorted(Comparator.comparing(DescriptionModel::getCreatedAt)).toList();
        
        if (!descriptions.isEmpty()) {
            buildSectionDescriptions(document, descriptions);
        }
    }

    private void buildSectionDescriptions(XWPFDocument document, List<DescriptionModel> descriptions) {
        if (document == null) throw new MyApplicationException("Document required");
        if (descriptions == null) throw new MyApplicationException("Descriptions required");

        List<DescriptionTemplateModel> descriptionTemplateModels = descriptions.stream().map(DescriptionModel::getDescriptionTemplate).toList();
        if (descriptionTemplateModels.isEmpty()) return;

        wordBuilder.addParagraphContent("Descriptions", document, ParagraphStyle.HEADER2, BigInteger.ZERO, 0);
//        for (DescriptionTemplateModel descriptionTemplateModelEntity : descriptionTemplateModels) {
//            XWPFParagraph templateParagraph = document.createParagraph();
//            XWPFRun runTemplateLabel = templateParagraph.createRun();
//            runTemplateLabel.setText("• " + descriptionTemplateModelEntity.getLabel());
//            runTemplateLabel.setColor("116a78");
//        }
        
        for (DescriptionModel descriptionModel : descriptions){
            buildSectionDescription(document, descriptionModel);
        }
    }

    private void buildSectionDescription(XWPFDocument document, DescriptionModel descriptionModel) {
        if (document == null) throw new MyApplicationException("Document required");
        if (descriptionModel == null) throw new MyApplicationException("DescriptionModel required");
        
        DescriptionTemplateModel descriptionTemplateModelFileModel = descriptionModel.getDescriptionTemplate();

        // Dataset Description custom style.
        XWPFParagraph datasetDescriptionParagraph = document.createParagraph();
        datasetDescriptionParagraph.setStyle("Heading4");
        datasetDescriptionParagraph.setSpacingBetween(1.5);
        XWPFRun datasetDescriptionRun = datasetDescriptionParagraph.createRun();
        datasetDescriptionRun.setText(descriptionModel.getLabel());
        datasetDescriptionRun.setFontSize(15);

        XWPFParagraph descriptionParagraph = document.createParagraph();
        wordBuilder.addParagraphContent(descriptionModel.getDescription(), document, ParagraphStyle.HTML, BigInteger.ZERO, 0);


        XWPFParagraph datasetTemplateParagraph = document.createParagraph();
        XWPFRun runDatasetTemplate1 = datasetTemplateParagraph.createRun();
        runDatasetTemplate1.setText("Template: ");
        runDatasetTemplate1.setColor("000000");
        XWPFRun runDatasetTemplate = datasetTemplateParagraph.createRun();
        runDatasetTemplate.setText(descriptionTemplateModelFileModel != null ? descriptionTemplateModelFileModel.getLabel() : "");
        runDatasetTemplate.setColor("116a78");


        XWPFParagraph datasetDescParagraph = document.createParagraph();
        XWPFRun runDatasetDescription1 = datasetDescParagraph.createRun();
        runDatasetDescription1.setText("Type: ");
        runDatasetDescription1.setColor("000000");
        XWPFRun runDatasetDescription = datasetDescParagraph.createRun();
        runDatasetDescription.setText(descriptionTemplateModelFileModel != null && descriptionTemplateModelFileModel.getType() != null ? descriptionTemplateModelFileModel.getType().getName() : "");
        runDatasetDescription.setColor("116a78");

        document.createParagraph();

        try {
            this.wordBuilder.build(document, descriptionModel.getDescriptionTemplate(), descriptionModel.getProperties(), new VisibilityServiceImpl(descriptionModel.getVisibilityStates()));
        } catch (Exception e) {
            logger.error(e.getMessage(), e);
        }
        // Page break at the end of the Dataset.
        XWPFParagraph parBreakDataset = document.createParagraph();
        parBreakDataset.setPageBreak(true);
    }


    private void buildPlanSectionField(PlanModel planEntity, XWPFDocument document, FieldModel fieldModel) throws InvalidApplicationException {
        if (fieldModel == null) throw new MyApplicationException("Field required");
        if (fieldModel.getCategory() == null) throw new MyApplicationException("Field is required" + fieldModel.getId() + " " + fieldModel.getLabel());
        switch (fieldModel.getCategory()){
            case System -> {
                buildPlanSectionSystemField(planEntity, document, (SystemFieldModel) fieldModel);
            }
            case Extra -> buildPlanSectionExtraField(planEntity, document, (ExtraFieldModel) fieldModel);
            case ReferenceType -> {
                buildPlanSectionReferenceTypeField(planEntity, document, (ReferenceTypeFieldModel) fieldModel);
            }
            case Upload -> {
                buildPlanSectionUploadField(planEntity, document, (UploadFieldModel) fieldModel);
            }
            default -> throw new MyApplicationException("Invalid type " + fieldModel.getCategory());
        }
    }

    private void buildPlanSectionReferenceTypeField(PlanModel planEntity, XWPFDocument document, ReferenceTypeFieldModel referenceField) {
        if (referenceField == null) throw new MyApplicationException("ReferenceField required");
        if (planEntity == null) throw new MyApplicationException("planEntity required");
        if (document == null) throw new MyApplicationException("Document required");
        if (referenceField.getReferenceType() == null) throw new MyApplicationException("ReferenceField type required");
        if (referenceField.getReferenceType().getCode() == null && !referenceField.getReferenceType().getCode().isBlank()) throw new IllegalArgumentException("ReferenceField type code required");

        XWPFParagraph systemFieldParagraph = document.createParagraph();
        XWPFRun runSyStemFieldTitle = systemFieldParagraph.createRun();
        runSyStemFieldTitle.setText(this.getReferenceFieldLabel(referenceField) + ": ");
        runSyStemFieldTitle.setColor("000000");

        List<ReferenceModel> referenceModels = this.getReferenceModelOfTypeCode(planEntity, referenceField.getReferenceType().getCode(), referenceField.getId());
        for (ReferenceModel reference : referenceModels) {
            XWPFRun runResearcher = systemFieldParagraph.createRun();
            if (this.wordFileTransformerServiceProperties.getResearcherReferenceCode().equalsIgnoreCase(referenceField.getReferenceType().getCode()) ||
                    this.wordFileTransformerServiceProperties.getOrganizationReferenceCode().equalsIgnoreCase(referenceField.getReferenceType().getCode())
            ) runResearcher.addBreak();
            if (this.wordFileTransformerServiceProperties.getLicenceReferenceCode().equalsIgnoreCase(referenceField.getReferenceType().getCode())) runResearcher.setText(reference.getReference());
            else runResearcher.setText(reference.getLabel());
            runResearcher.setColor("116a78");
        }
    }

    private void buildPlanSectionUploadField(PlanModel planEntity, XWPFDocument document, UploadFieldModel uploadFieldModel) {
        if (uploadFieldModel == null) throw new MyApplicationException("UploadFieldModel required");

        XWPFParagraph uploadFieldParagraph = document.createParagraph();
        uploadFieldParagraph.setSpacingBetween(1.0);
        XWPFRun runUploadFieldLabel = uploadFieldParagraph.createRun();
        runUploadFieldLabel.setText(uploadFieldModel.getLabel() + ": ");
        runUploadFieldLabel.setColor("000000");

        PlanBlueprintValueModel planBlueprintValueModel = planEntity.getProperties() != null && planEntity.getProperties().getPlanBlueprintValues() != null ? planEntity.getProperties().getPlanBlueprintValues().stream().filter(x -> uploadFieldModel.getId().equals(x.getFieldId())).findFirst().orElse(null) : null;
        if (planBlueprintValueModel != null && planBlueprintValueModel.getValue() != null && !planBlueprintValueModel.getValue().isBlank()) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.setPageBreak(true);
            paragraph.setSpacingAfter(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER); //GK: Center the image if it is too small
            XWPFRun run = paragraph.createRun();
            FileEnvelopeModel itemTyped = planBlueprintValueModel.getFile();
            if (itemTyped == null) return;
            try {

                String fileName = itemTyped.getFilename();
                String fileType = itemTyped.getMimeType();
                if (IMAGE_TYPE_MAP.containsKey(fileType)) {
                    int format;
                    format = IMAGE_TYPE_MAP.getOrDefault(fileType, 0);
                    byte[] file;
                    if (this.wordFileTransformerServiceProperties.isUseSharedStorage() && itemTyped.getFileRef() != null && !itemTyped.getFileRef().isBlank()) {
                        file = this.fileStorageService.readFile(itemTyped.getFileRef());
                    } else {
                        file = itemTyped.getFile();
                    }
                    InputStream image = new ByteArrayInputStream(file);
                    ImageInputStream iis = ImageIO.createImageInputStream(new ByteArrayInputStream(file));
                    Iterator<ImageReader> readers = ImageIO.getImageReaders(iis);
                    if (readers.hasNext()) {
                        ImageReader reader = readers.next();
                        reader.setInput(iis);

                        int initialImageWidth = reader.getWidth(0);
                        int initialImageHeight = reader.getHeight(0);

                        float ratio = initialImageHeight / (float) initialImageWidth;

                        int marginLeftInDXA = this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgMar().getLeft());
                        int marginRightInDXA = this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgMar().getRight());
                        int pageWidthInDXA = this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgSz().getW());
                        int pageWidth = Math.round((pageWidthInDXA - marginLeftInDXA - marginRightInDXA) / (float) 20); // /20 converts dxa to points

                        int imageWidth = Math.round(initialImageWidth * (float) 0.75);    // *0.75 converts pixels to points
                        int width = Math.min(imageWidth, pageWidth);

                        int marginTopInDXA =  this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgMar().getTop());
                        int marginBottomInDXA = this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgMar().getBottom());
                        int pageHeightInDXA = this.toIntFormBigInteger(document.getDocument().getBody().getSectPr().getPgSz().getH());
                        int pageHeight = Math.round((pageHeightInDXA - marginTopInDXA - marginBottomInDXA) / (float) 20);    // /20 converts dxa to points

                        int imageHeight = Math.round(initialImageHeight * ((float) 0.75));  // *0.75 converts pixels to points

                        int height = Math.round(width * ratio);
                        if (height > pageHeight) {
                            // height calculated with ratio is too large. Image may have Portrait (vertical) orientation. Recalculate image dimensions.
                            height = Math.min(imageHeight, pageHeight);
                            width = Math.round(height / ratio);
                        }

                        run.addPicture(image, format, fileName, Units.toEMU(width), Units.toEMU(height));
                        paragraph.setPageBreak(false);
                        imageCount++;
                        XWPFParagraph captionParagraph = document.createParagraph();
                        captionParagraph.setAlignment(ParagraphAlignment.CENTER);
                        captionParagraph.setSpacingBefore(0);
                        captionParagraph.setStyle("Caption");
                        XWPFRun captionRun = captionParagraph.createRun();
                        captionRun.setText("Image " + imageCount);
                    }
                } else {
                    if(planBlueprintValueModel.getFile() != null && planBlueprintValueModel.getFile().getFilename() != null && !planBlueprintValueModel.getFile().getFilename().isBlank()) {
                        XWPFRun runUploadFieldInput = uploadFieldParagraph.createRun();
                        runUploadFieldInput.setText(planBlueprintValueModel.getFile().getFilename());
                        runUploadFieldInput.setColor("116a78");
                    }
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
            }
        }
    }

    private int toIntFormBigInteger(Object object){
        try {
            if (object instanceof BigInteger) return ((BigInteger) object).intValue();
            return (int) object;
        } catch (Exception e){
            logger.error(e.getMessage(), e);
            return 0;
        }
    }

    private String getReferenceFieldLabel(ReferenceTypeFieldModel referenceTypeField) {
        if (referenceTypeField == null) return "";
        if (referenceTypeField.getLabel() != null && !referenceTypeField.getLabel().isBlank()) return referenceTypeField.getLabel();

        return referenceTypeField.getReferenceType().getName();
    }

    private void buildPlanSectionSystemField(PlanModel planEntity, XWPFDocument document, SystemFieldModel systemField) {
        if (systemField == null) throw new MyApplicationException("SystemField required");
        if (planEntity == null) throw new MyApplicationException("planEntity required");
        if (document == null) throw new MyApplicationException("Document required");

        if (PlanBlueprintSystemFieldType.Language.equals(systemField.getSystemFieldType()) || PlanBlueprintSystemFieldType.User.equals(systemField.getSystemFieldType())) return;


        XWPFParagraph systemFieldParagraph = document.createParagraph();
        XWPFRun runSyStemFieldTitle = systemFieldParagraph.createRun();
        runSyStemFieldTitle.setText(this.getSystemFieldLabel(systemField) + ": ");
        runSyStemFieldTitle.setColor("000000");
        
        switch (systemField.getSystemFieldType()) {
            case Title:
                XWPFRun runTitle = systemFieldParagraph.createRun();
                runTitle.setText(planEntity.getLabel());
                runTitle.setColor("116a78");
                break;
            case Description:
                wordBuilder.addParagraphContent(planEntity.getDescription(), document, ParagraphStyle.HTML, BigInteger.ZERO, 0);
                break;
            case AccessRights:
                if (planEntity.getAccessType() != null) {
                    XWPFRun runAccessRights = systemFieldParagraph.createRun();
                    runAccessRights.setText(planEntity.getAccessType().equals(PlanAccessType.Public) ? "Public" : "Restricted"); //TODO
                    runAccessRights.setColor("116a78");
                }
                break;
            case Contact:
                List<String> contacts = new ArrayList<>();
                
                if (planEntity.getProperties() != null && planEntity.getProperties().getContacts() != null && !planEntity.getProperties().getContacts().isEmpty()) {
                    for (PlanContactModel contactModel : planEntity.getProperties().getContacts()){
                        String contact;
                        contact = (contactModel.getLastName() == null ? "" : contactModel.getLastName()) + " " + (contactModel.getFirstName() == null ? "" : contactModel.getFirstName());
                        if (contactModel.getEmail() != null && !contactModel.getEmail().isEmpty()) contact = contact + " (" + contactModel.getEmail() +")";
                        contacts.add(contact.trim());
                    }
                } 
                
                if (!contacts.isEmpty()) {
                    XWPFRun runContact = systemFieldParagraph.createRun();
                    runContact.setText(String.join(", ", contacts));
                    runContact.setColor("116a78");
                }
                break;
            case User:
            case Language:
                break;
            default:
                throw new MyApplicationException("Invalid type " + systemField.getSystemFieldType());
        }
    }
    
    private String getSystemFieldLabel(SystemFieldModel systemField) {
        if (systemField == null) return "";
        if (systemField.getLabel() != null && !systemField.getLabel().isBlank()) return systemField.getLabel();

	    return switch (systemField.getSystemFieldType()) {
		    case Title -> this.messageSource.getMessage("SystemField_Title_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    case Description -> this.messageSource.getMessage("SystemField_Description_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    case AccessRights -> this.messageSource.getMessage("SystemField_AccessRights_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    case Contact -> this.messageSource.getMessage("SystemField_Contact_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    case User -> this.messageSource.getMessage("SystemField_User_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    case Language -> this.messageSource.getMessage("SystemField_Language_Label", new Object[]{}, LocaleContextHolder.getLocale());
		    default -> throw new MyApplicationException("Invalid type " + systemField.getSystemFieldType());
	    };
    }

    private void buildPlanSectionExtraField(PlanModel planEntity, XWPFDocument document, ExtraFieldModel extraFieldModel) {
        if (extraFieldModel == null) throw new MyApplicationException("ExtraFieldModel required");
        XWPFParagraph extraFieldParagraph = document.createParagraph();
        extraFieldParagraph.setSpacingBetween(1.0);
        XWPFRun runExtraFieldLabel = extraFieldParagraph.createRun();
        runExtraFieldLabel.setText(extraFieldModel.getLabel() + ": ");
        runExtraFieldLabel.setColor("000000");

        XWPFRun runExtraFieldInput = extraFieldParagraph.createRun();
        PlanBlueprintValueModel planBlueprintValueModel = planEntity.getProperties() != null && planEntity.getProperties().getPlanBlueprintValues() != null ? planEntity.getProperties().getPlanBlueprintValues().stream().filter(x -> extraFieldModel.getId().equals(x.getFieldId())).findFirst().orElse(null) : null;
        if (planBlueprintValueModel != null) {
            switch (extraFieldModel.getDataType()) {
                case RichTex:
                    if(planBlueprintValueModel.getValue() != null && !planBlueprintValueModel.getValue().isBlank()) wordBuilder.addParagraphContent(planBlueprintValueModel.getValue(), document, ParagraphStyle.HTML, BigInteger.ZERO, 0);
                    break;
                case Number:
                    if(planBlueprintValueModel.getNumberValue() != null) {
                        runExtraFieldInput.setText(DecimalFormat.getNumberInstance().format(planBlueprintValueModel.getNumberValue()));
                        runExtraFieldInput.setColor("116a78");
                    }
                    break;
                case Date:
                    if(planBlueprintValueModel.getDateValue() != null){
                        runExtraFieldInput.setText(DateTimeFormatter.ofPattern("yyyy-MM-dd").withZone(ZoneId.systemDefault()).format(planBlueprintValueModel.getDateValue()));
                        runExtraFieldInput.setColor("116a78");
                    }
                    break;
                case Text:
                    if(planBlueprintValueModel.getValue() != null && !planBlueprintValueModel.getValue().isBlank()) {
                        runExtraFieldInput.setText(planBlueprintValueModel.getValue());
                        runExtraFieldInput.setColor("116a78");
                    }
                    break;
                default:
                    throw new MyApplicationException("Invalid type " + extraFieldModel.getDataType());
            }
        }
    }

    private String getPlanFileName(PlanModel planModel, String extension){
        if (planModel == null) throw new MyApplicationException("PlanEntity required");

        List<ReferenceModel> grants = this.getReferenceModelOfTypeCode(planModel, this.wordFileTransformerServiceProperties.getGrantReferenceCode(), null);
        String fileName = null;
        if (planModel.getLabel() != null){
            return planModel.getLabel() + extension;
        }
        if (!grants.isEmpty() && grants.getFirst().getLabel() != null) {
            fileName = "PLAN_" + grants.getFirst().getLabel();
            fileName += "_" + planModel.getVersion();

        }
       
        return fileName + extension;
    }

    private byte[] buildDescriptionWordDocument(DescriptionModel descriptionModel) throws IOException {
        if (descriptionModel == null) throw new MyApplicationException("DescriptionEntity required");
        PlanModel planEntity = descriptionModel.getPlan();
        if (planEntity == null)  throw new MyApplicationException("plan is invalid");

        XWPFDocument document = null;
        if (descriptionModel.getDescriptionTemplate() != null && descriptionModel.getDescriptionTemplate().getDefinition().getPlugins() != null && !descriptionModel.getDescriptionTemplate().getDefinition().getPlugins().isEmpty()) {
            document = this.getCustomDocument(descriptionModel.getDescriptionTemplate().getDefinition().getPlugins(), PluginEntityType.Description);
        }
        if (document == null) {
            try {
                Resource resource = resourceLoader.getResource(this.wordFileTransformerServiceProperties.getWordDescriptionTemplate());
                try(InputStream inputStream = resource.getInputStream()) {
                    document = new XWPFDocument(inputStream);
                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
                throw new RuntimeException(e);
            }
        }

        this.wordBuilder.fillFirstPage(planEntity, descriptionModel, document, true);
        this.wordBuilder.fillFooter(planEntity, descriptionModel, document);
        this.wordBuilder.fillHeader(planEntity, descriptionModel, document);

        int powered_pos = this.wordBuilder.findPosOfPoweredBy(document);
        XWPFParagraph powered_par = null;
        XWPFParagraph argos_img_par = null;
        if(powered_pos != -1) {
            powered_par = document.getParagraphArray(powered_pos);
            argos_img_par = document.getParagraphArray(powered_pos + 1);
        }

        this.wordBuilder.build(document, descriptionModel.getDescriptionTemplate(), descriptionModel.getProperties(), new VisibilityServiceImpl(descriptionModel.getVisibilityStates()));
        
        if(powered_pos != -1) {
            document.getLastParagraph().setPageBreak(false);
            document.createParagraph();
            document.setParagraph(powered_par, document.getParagraphs().size() - 1);

            document.createParagraph();
            document.setParagraph(argos_img_par, document.getParagraphs().size() - 1);

            document.removeBodyElement(powered_pos + 1);
            document.removeBodyElement(powered_pos + 1);
        }
        
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        document.write(out);
        byte[] bytes = out.toByteArray();
        out.close();

        return bytes;
    }

    private String getDescriptionFileName(DescriptionModel descriptionModel, String extension){
        if (descriptionModel == null) throw new MyApplicationException("DescriptionEntity required");
        String fileName = descriptionModel.getLabel().replaceAll("[^a-zA-Z0-9+ ]", "");
        
        return fileName + extension;
    }

    private XWPFDocument getCustomDocument(List<PluginModel> plugins, PluginEntityType entityType) {
        try {
            if (plugins != null && !plugins.isEmpty()) {
                PluginModel plugin = plugins.stream().filter(x -> x.getCode().equals(this.wordFileTransformerServiceProperties.getTransformerId()) && x.getType().equals(PluginType.FileTransformer)).findFirst().orElse(null);
                if (plugin != null && plugin.getFields() != null) {
                    if (this.wordFileTransformerServiceProperties.getConfigurationFields() != null) {
                        List<ConfigurationField> filteredConfigurationFields = this.wordFileTransformerServiceProperties.getConfigurationFields().stream().filter(x -> x.getAppliesTo() != null && x.getAppliesTo().contains(entityType)).toList();
                        if (!filteredConfigurationFields.isEmpty()) {
                            PluginFieldModel field = plugin.getFields().stream().filter(y -> y.getFile() != null && filteredConfigurationFields.stream().map(ConfigurationField::getCode).toList().contains(y.getCode())).findFirst().orElse(null);
                            if (field != null && field.getFile() != null && field.getFile().getFile() != null) {
                                return new XWPFDocument(new ByteArrayInputStream(field.getFile().getFile()));
                            }
                        }
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
