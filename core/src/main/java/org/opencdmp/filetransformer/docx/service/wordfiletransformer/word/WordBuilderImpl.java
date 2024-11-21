package org.opencdmp.filetransformer.docx.service.wordfiletransformer.word;

import gr.cite.tools.exception.MyApplicationException;
import org.opencdmp.commonmodels.enums.FieldType;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.description.PropertyDefinitionFieldSetItemModel;
import org.opencdmp.commonmodels.models.description.PropertyDefinitionFieldSetModel;
import org.opencdmp.commonmodels.models.description.PropertyDefinitionModel;
import org.opencdmp.commonmodels.models.descriptiotemplate.*;
import org.opencdmp.commonmodels.models.descriptiotemplate.fielddata.*;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.planreference.PlanReferenceModel;
import org.opencdmp.commonmodels.models.reference.ReferenceFieldModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageServiceProperties;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.opencdmp.filetransformer.docx.model.PidLink;
import org.opencdmp.filetransformer.docx.model.interfaces.ApplierWithValue;
import org.opencdmp.filetransformer.docx.service.pid.PidService;
import org.opencdmp.filetransformer.docx.model.enums.ParagraphStyle;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.visibility.VisibilityService;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.NodeTraversor;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerConfiguration;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.config.ConfigurableBeanFactory;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import javax.management.InvalidApplicationException;
import java.io.*;
import java.math.BigInteger;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

import static org.apache.poi.xwpf.usermodel.Document.*;

@Component
@Scope(value = ConfigurableBeanFactory.SCOPE_PROTOTYPE)
public class WordBuilderImpl implements WordBuilder {
    private static final Logger logger = LoggerFactory.getLogger(WordBuilderImpl.class);
    private static final Map<String, Integer> IMAGE_TYPE_MAP = Map.of(
            "image/jpeg", PICTURE_TYPE_JPEG,
            "image/png", PICTURE_TYPE_PNG,
            "image/gif", PICTURE_TYPE_GIF,
            "image/tiff", PICTURE_TYPE_TIFF,
            "image/bmp", PICTURE_TYPE_BMP,
            "image/wmf", PICTURE_TYPE_WMF
    );
    private BigInteger numId;
    private Integer indent;
    private Integer imageCount;
    private final CTAbstractNum cTAbstractNum;
    private final FileStorageService fileStorageService;
    private final WordFileTransformerServiceProperties wordFileTransformerServiceProperties;
    private final PidService pidService;
    private final Map<ParagraphStyle, ApplierWithValue<XWPFDocument, Object, XWPFParagraph>> options = new HashMap<>();
    private final Map<ParagraphStyle, ApplierWithValue<XWPFTableCell, Object, XWPFParagraph>> optionsInTable = new HashMap<>();

    public WordBuilderImpl(FileStorageService fileStorageService, WordFileTransformerServiceProperties wordFileTransformerServiceProperties, PidService pidService) {
	    this.fileStorageService = fileStorageService;
	    this.wordFileTransformerServiceProperties = wordFileTransformerServiceProperties;
        this.pidService = pidService;
	    this.cTAbstractNum = CTAbstractNum.Factory.newInstance();
        this.cTAbstractNum.setAbstractNumId(BigInteger.valueOf(1));
        this.indent = 0;
        this.imageCount = 0;
        this.buildOptions();
        this.buildOptionsInTable();
    }

    private void buildOptionsInTable() {
        this.optionsInTable.put(ParagraphStyle.TEXT, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.addParagraph();
            XWPFRun run = paragraph.createRun();
            if (item != null)
                run.setText("" + item);
            run.setFontSize(11);
            return paragraph;
        });
        this.optionsInTable.put(ParagraphStyle.HTML, (mainDocumentPart, item) -> {
            Document htmlDoc = Jsoup.parse(((String) item).replaceAll("\n", "<br>"));
            HtmlToWorldBuilder htmlToWorldBuilder = HtmlToWorldBuilder.convertInTable(mainDocumentPart, htmlDoc, 0);
            return htmlToWorldBuilder.getParagraph();
        });
        this.optionsInTable.put(ParagraphStyle.TITLE, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.addParagraph();
            paragraph.setStyle("Title");
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText((String) item);
            run.setBold(true);
            run.setFontSize(14);
            return paragraph;
        });
        this.optionsInTable.put(ParagraphStyle.IMAGE, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.addParagraph();
            XWPFRun run = paragraph.createRun();
            if (item instanceof FileEnvelopeModel)
                run.setText(((FileEnvelopeModel)item).getFilename());
            run.setFontSize(11);
            run.setItalic(true);
            return paragraph;
        });
    }

    private void buildOptions() {
        this.options.put(ParagraphStyle.TEXT, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            XWPFRun run = paragraph.createRun();
            if (item != null)
                run.setText("" + item);
            run.setFontSize(11);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HTML, (mainDocumentPart, item) -> {
            Document htmlDoc = Jsoup.parse(((String) item).replaceAll("\n", "<br>"));
            HtmlToWorldBuilder htmlToWorldBuilder = HtmlToWorldBuilder.convert(mainDocumentPart, htmlDoc, this.indent);
            return htmlToWorldBuilder.getParagraph();
        });
        this.options.put(ParagraphStyle.TITLE, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Title");
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun run = paragraph.createRun();
            run.setText((String) item);
            run.setBold(true);
            run.setFontSize(14);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER1, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading1");
            XWPFRun run = paragraph.createRun();
            run.setText((String) item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER2, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading2");
            XWPFRun run = paragraph.createRun();
            run.setText("" + item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER3, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading3");
            XWPFRun run = paragraph.createRun();
            run.setText("" + item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER4, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading4");
            XWPFRun run = paragraph.createRun();
            run.setText((String) item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER5, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading5");
            XWPFRun run = paragraph.createRun();
            run.setText("" + item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.HEADER6, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setStyle("Heading6");
            XWPFRun run = paragraph.createRun();
            run.setText("" + item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.FOOTER, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText((String) item);
            return paragraph;
        });
        this.options.put(ParagraphStyle.COMMENT, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("" + item);
            run.setItalic(true);
            return paragraph;
        });
        this.options.put(ParagraphStyle.IMAGE, (mainDocumentPart, item) -> {
            XWPFParagraph paragraph = mainDocumentPart.createParagraph();
            paragraph.setPageBreak(true);
            paragraph.setSpacingAfter(0);
            paragraph.setAlignment(ParagraphAlignment.CENTER); //GK: Center the image if it is too small
            XWPFRun run = paragraph.createRun();
            FileEnvelopeModel itemTyped = (FileEnvelopeModel)item;
            if (itemTyped == null) return paragraph;
            try {

                String fileName = itemTyped.getFilename();
                String fileType = itemTyped.getMimeType();
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

                    int marginLeftInDXA = this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgMar().getLeft());
                    int marginRightInDXA = this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgMar().getRight());
                    int pageWidthInDXA = this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgSz().getW());
                    int pageWidth = Math.round((pageWidthInDXA - marginLeftInDXA - marginRightInDXA) / (float) 20); // /20 converts dxa to points

                    int imageWidth = Math.round(initialImageWidth * (float) 0.75);    // *0.75 converts pixels to points
                    int width = Math.min(imageWidth, pageWidth);

                    int marginTopInDXA =  this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgMar().getTop());
                    int marginBottomInDXA = this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgMar().getBottom());
                    int pageHeightInDXA = this.toIntFormBigInteger(mainDocumentPart.getDocument().getBody().getSectPr().getPgSz().getH());
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
                    XWPFParagraph captionParagraph = mainDocumentPart.createParagraph();
                    captionParagraph.setAlignment(ParagraphAlignment.CENTER);
                    captionParagraph.setSpacingBefore(0);
                    captionParagraph.setStyle("Caption");
                    XWPFRun captionRun = captionParagraph.createRun();
                    captionRun.setText("Image " + imageCount);

                }
            } catch (Exception e) {
                logger.error(e.getMessage(), e);
            }
            return paragraph;
        });
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

    @Override
    public void build(XWPFDocument document, DescriptionTemplateModel descriptionTemplate, PropertyDefinitionModel propertyDefinitionModel, VisibilityService visibilityService) {
        createPages(descriptionTemplate.getDefinition().getPages(), propertyDefinitionModel, document, visibilityService);
    }

    private void createPages(List<PageModel> datasetProfilePages, PropertyDefinitionModel propertyDefinitionModel, XWPFDocument mainDocumentPart, VisibilityService visibilityService) {
	    for (PageModel item : datasetProfilePages) {
                if (item.getSections() != null) {
                    try {
                        XWPFParagraph paragraph = addParagraphContent(item.getOrdinal() + 1 + " " + item.getTitle(), mainDocumentPart, ParagraphStyle.HEADER5, numId, 0);
                        mainDocumentPart.getPosOfParagraph(paragraph);
                        if (visibilityService.isVisible(item.getId(), null)) {
                            createSections(item.getSections(), propertyDefinitionModel, mainDocumentPart, 1, false, item.getOrdinal() + 1, null, visibilityService);
                        }
                    } catch (Exception e) {
                        logger.error(e.getMessage(), e);
                    }
                }
            }
    }

    private boolean createSections(List<SectionModel> sections, PropertyDefinitionModel propertyDefinitionModel, XWPFDocument mainDocumentPart, Integer indent, Boolean createListing, Integer page, String sectionString, VisibilityService visibilityService) {
        if (createListing) this.addListing(indent, false, true);
        boolean hasAnySectionValue = false;
        
        for (SectionModel section : sections) {
            if (!visibilityService.isVisible(section.getId(), null)) continue;
            boolean hasValue = false;
            int paragraphPos = -1;
            String tempSectionString = sectionString != null ? sectionString + "." + (section.getOrdinal() + 1) : "" + (section.getOrdinal() + 1);
            if (!createListing) {
                XWPFParagraph paragraph = addParagraphContent(page + "." + tempSectionString + " " + section.getTitle(), mainDocumentPart, ParagraphStyle.HEADER5, numId, indent);
                paragraphPos = mainDocumentPart.getPosOfParagraph(paragraph);
            }
            if (section.getSections() != null) {
                hasValue = createSections(section.getSections(), propertyDefinitionModel, mainDocumentPart, indent + 1, createListing, page, tempSectionString, visibilityService);
            }
            if (section.getFieldSets() != null) {
                hasValue = createFieldSetFields(section.getFieldSets(), propertyDefinitionModel, mainDocumentPart, indent + 1, createListing, page, tempSectionString, visibilityService);
            }

            if (!hasValue && paragraphPos > -1) {
                mainDocumentPart.removeBodyElement(paragraphPos);
            }
            hasAnySectionValue = hasAnySectionValue || hasValue;
        }
        
        return hasAnySectionValue;
    }


    private Boolean createFieldSetFields(List<FieldSetModel> fieldSetModels, PropertyDefinitionModel propertyDefinitionModel, XWPFDocument mainDocumentPart, Integer indent, Boolean createListing, Integer page, String section, VisibilityService visibilityService) {
        if (createListing) this.addListing(indent, true, true);
        boolean hasValue = false;
        boolean returnedValue = false;

        for (FieldSetModel fieldSetModel : fieldSetModels) {
            PropertyDefinitionFieldSetModel propertyDefinitionFieldSetModel = propertyDefinitionModel.getFieldSets().getOrDefault(fieldSetModel.getId(), null);
            List<PropertyDefinitionFieldSetItemModel> propertyDefinitionFieldSetItemModels = propertyDefinitionFieldSetModel != null && propertyDefinitionFieldSetModel.getItems() != null ? propertyDefinitionFieldSetModel.getItems() : new ArrayList<>();
            propertyDefinitionFieldSetItemModels = propertyDefinitionFieldSetItemModels.stream().sorted(Comparator.comparingInt(PropertyDefinitionFieldSetItemModel::getOrdinal)).toList();
            if (propertyDefinitionFieldSetItemModels.stream().anyMatch(x -> visibilityService.isVisible(fieldSetModel.getId(), x.getOrdinal()))) {

                char c = 'a';
                int multiplicityItems = 0;
                boolean hasMultiplicityItems = false;
                int paragraphPos = -1;
                int paragraphPosInner = -1;
                if (fieldSetModel.getTitle() != null && !fieldSetModel.getTitle().isEmpty() && !createListing) {
                    XWPFParagraph paragraph = addParagraphContent(page + "." + section + "." + (fieldSetModel.getOrdinal() + 1) + " " + fieldSetModel.getTitle(), mainDocumentPart, ParagraphStyle.HEADER6, numId, indent);
//                    CTDecimalNumber number = paragraph.getCTP().getPPr().getNumPr().addNewIlvl();
//                    number.setVal(BigInteger.valueOf(indent));
                    paragraphPos = mainDocumentPart.getPosOfParagraph(paragraph);
                    if (fieldSetModel.getMultiplicity() != null && !fieldSetModel.getMultiplicity().getTableView() && propertyDefinitionFieldSetItemModels.size() > 1) {
                        XWPFParagraph paragraphInner = addParagraphContent(c + ". ", mainDocumentPart, ParagraphStyle.TEXT, numId, indent);
                        paragraphPosInner = mainDocumentPart.getPosOfParagraph(paragraphInner);
                        hasMultiplicityItems = true;
                        multiplicityItems++;
                    }
                }
                XWPFTable tbl = null;
                XWPFTableRow row = null;
                int numOfRows = 0;
                if (fieldSetModel.getMultiplicity() != null && fieldSetModel.getMultiplicity().getTableView()) {
                    tbl = mainDocumentPart.createTable();
                    tbl.setTableAlignment(TableRowAlign.CENTER);
                    mainDocumentPart.createParagraph();
                    createHeadersInTable(fieldSetModel.getFields(), propertyDefinitionFieldSetItemModels.getFirst(), tbl, visibilityService);
                    numOfRows = tbl.getRows().size();
                    row = tbl.createRow();
                }
                if (fieldSetModel.getMultiplicity() != null && fieldSetModel.getMultiplicity().getTableView()) {
                    hasValue = createFieldsInTable(fieldSetModel, propertyDefinitionFieldSetItemModels.getFirst(), row, indent, createListing, hasMultiplicityItems, numOfRows, visibilityService);
                    if (!hasValue && propertyDefinitionFieldSetItemModels.size() > 1 && tbl != null) {
                        tbl.removeRow(numOfRows);
                    } else if (!hasValue && tbl != null) {
                        for (int i = numOfRows; i >= 0; i--) {
                            tbl.removeRow(i);
                        }
                    } else numOfRows++;
                } else {
                    hasValue = createFields(fieldSetModel, propertyDefinitionFieldSetItemModels.getFirst(), mainDocumentPart, indent, createListing, hasMultiplicityItems, visibilityService);
                }
                if (hasValue) {
                    returnedValue = true;
                } else if (paragraphPosInner > -1) {
                    mainDocumentPart.removeBodyElement(paragraphPosInner);
                    c--;
                    multiplicityItems--;
                }
                if (propertyDefinitionFieldSetItemModels.size() > 1) {
                    int fieldsCount = 0;
                    for (PropertyDefinitionFieldSetItemModel multiplicityFieldset : propertyDefinitionFieldSetItemModels.stream().skip(1).toList()) {
                        paragraphPosInner = -1;
                        if (fieldSetModel.getMultiplicity() != null && !fieldSetModel.getMultiplicity().getTableView() && !createListing) {
                            c++;
//                            addParagraphContent(c + ". ", mainDocumentPart, ParagraphStyle.HEADER6, numId);
                            XWPFParagraph paragraphInner = addParagraphContent(c + ". ", mainDocumentPart, ParagraphStyle.TEXT, numId, indent);
                            paragraphPosInner = mainDocumentPart.getPosOfParagraph(paragraphInner);
                            hasMultiplicityItems = true;
                            multiplicityItems++;
                        }
//                        hasValue = createFields(multiplicityFieldset.getFields(), mainDocumentPart, 3, createListing, visibilityRuleService, hasMultiplicityItems);
                        boolean hasValueInner = false;
                        if (fieldSetModel.getMultiplicity() != null && fieldSetModel.getMultiplicity().getTableView() && tbl != null) {
                            row = tbl.createRow();
                            hasValueInner = createFieldsInTable(fieldSetModel, multiplicityFieldset, row, indent, createListing, hasMultiplicityItems, numOfRows, visibilityService);
                            if (!hasValueInner && numOfRows <= 1 && fieldsCount == propertyDefinitionFieldSetItemModels.size()-2) { //-2 because we skip 1
                                for (int i = numOfRows; i >= 0; i--) {
                                    tbl.removeRow(i);
                                }
                            } else if (!hasValueInner) {
                                tbl.removeRow(numOfRows);
                            } else numOfRows++;
                        } else {
                            hasValueInner = createFields(fieldSetModel, multiplicityFieldset, mainDocumentPart, indent, createListing, hasMultiplicityItems, visibilityService);
                        }
//                        if(hasValue){
                        if (hasValueInner) {
                            hasValue = true;
                            returnedValue = true;
                        } else if (paragraphPosInner > -1) {
                            mainDocumentPart.removeBodyElement(paragraphPosInner);
                            c--;
                            multiplicityItems--;
                        }

                        fieldsCount++;
                    }
                    if (multiplicityItems == 1) {
                        String text = mainDocumentPart.getLastParagraph().getRuns().getFirst().getText(0);
                        if (text.equals("a. ")) {
                            mainDocumentPart.getLastParagraph().removeRun(0);
                        }
                    }
                }
                if (propertyDefinitionFieldSetModel.getComment() != null && !propertyDefinitionFieldSetModel.getComment().isEmpty()) {
                    addParagraphContent("<i>Comment:</i>\n" + propertyDefinitionFieldSetModel.getComment(), mainDocumentPart, ParagraphStyle.HTML, numId, indent);
                    hasValue = true;
                    returnedValue = true;
                }
                if (!hasValue && paragraphPos > -1) {
                    mainDocumentPart.removeBodyElement(paragraphPos);
                }
            }
        }

        return returnedValue;
    }

    private void createHeadersInTable(List<FieldModel> fields, PropertyDefinitionFieldSetItemModel propertyDefinitionFieldSetItemModel, XWPFTable table, VisibilityService visibilityService) {
        boolean atLeastOneHeader = false;
        List<FieldModel> tempFields = fields.stream().sorted(Comparator.comparingInt(FieldModel::getOrdinal)).toList();
        int index = 0;
        XWPFTableRow row = table.getRow(0);
        for (FieldModel field : tempFields) {
            if (field.getIncludeInExport() && visibilityService.isVisible(field.getId(), propertyDefinitionFieldSetItemModel.getOrdinal())) {
                XWPFTableCell cell;
                if (index == 0) {
                    cell = row.getCell(0);
                } else {
                    cell = row.createCell();
                }
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf("CENTER"));
                String label = field.getData().getLabel();

                if (label.isEmpty()) {
                    FieldType fieldType = field.getData().getFieldType();
                    if (fieldType.equals(FieldType.REFERENCE_TYPES)) {
                        ReferenceTypeDataModel referenceTypeDataModel = (ReferenceTypeDataModel) field.getData();
                        label = referenceTypeDataModel.getReferenceType().getName();
                    }
                }

                if (label != null && !label.isBlank()) {
                    XWPFParagraph paragraph = cell.getParagraphs().getFirst();
                    paragraph.setIndentationFirstLine(50);
                    XWPFRun run = paragraph.createRun();
                    run.setText(label);
                    run.setBold(true);
                    run.setFontSize(12);
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                    paragraph.setSpacingBefore(100);

                    atLeastOneHeader = true;
                }
            }
            index++;
        }

        if (!atLeastOneHeader) {
            table.removeRow(0);
        }
    }

    private Boolean createFieldsInTable(FieldSetModel fieldSetModel, PropertyDefinitionFieldSetItemModel propertyDefinitionFieldSetItemModel, XWPFTableRow mainDocumentPart,
                                        Integer indent, Boolean createListing, boolean hasMultiplicityItems, int numOfRows, VisibilityService visibilityService) {
        int numOfCells = 0;
        boolean hasValue = false;
        List<FieldModel> tempFields = fieldSetModel.getFields().stream().sorted(Comparator.comparingInt(FieldModel::getOrdinal)).toList();
        for (FieldModel field : tempFields) {
            if (field.getIncludeInExport() && visibilityService.isVisible(field.getId(), propertyDefinitionFieldSetItemModel.getOrdinal())) {
                if (!createListing) {
                    org.opencdmp.commonmodels.models.description.FieldModel fieldValueModel = propertyDefinitionFieldSetItemModel.getFields().getOrDefault(field.getId(), null);
                    if (field.getData().getFieldType().equals(FieldType.UPLOAD)) {
                        boolean isImage = false;
                        for (UploadDataModel.UploadOptionModel type : ((UploadDataModel) field.getData()).getTypes()) {
                            String fileFormat = type.getValue();
                            if (IMAGE_TYPE_MAP.containsKey(fileFormat)) {
                                isImage = true;
                                break;
                            }
                        }
                        if (isImage) {
                            if (fieldValueModel != null && fieldValueModel.getTextValue() != null && !fieldValueModel.getTextValue().isEmpty()) {
                                XWPFParagraph paragraph = addCellContent(fieldValueModel.getFile(), mainDocumentPart, ParagraphStyle.IMAGE, numId, 0, numOfRows, numOfCells, 0);
                                if (paragraph != null) {
                                    hasValue = true;
                                }
                                if (hasMultiplicityItems) {
                                    hasMultiplicityItems = false;
                                }
                            }
                        }
                    } else if (fieldValueModel != null) {
                        this.indent = indent;
                        boolean isResearcher = false;
                        if (field.getData() instanceof ReferenceTypeDataModel) {
                            isResearcher = ((ReferenceTypeDataModel) field.getData()).getReferenceType().getCode().equals(this.wordFileTransformerServiceProperties.getResearcherReferenceCode());
                        }

                        List<String> extractValues = this.extractValues(field, fieldValueModel);
                        if (!extractValues.isEmpty()){
                            int numOfValuesInCell = 0;
                            for (String extractValue : extractValues){
                                boolean orcidResearcher = false;
                                String orcId = null;
                                if (isResearcher && extractValue.contains("orcid:")) {
                                    orcId = extractValue.substring(extractValue.indexOf(':') + 1, extractValue.indexOf(')'));
                                    extractValue = extractValue.substring(0, extractValue.indexOf(':') + 1) + " ";
                                    orcidResearcher = true;
                                }
                                if (extractValues.size() > 1) extractValue = "• " + extractValue;
                                if (hasMultiplicityItems) {
                                    XWPFParagraph paragraph = mainDocumentPart.getCell(mainDocumentPart.getTableCells().size()).addParagraph();
                                    paragraph.createRun().setText(extractValue);
                                    if (orcidResearcher) {
                                        XWPFHyperlinkRun run = paragraph.createHyperlinkRun("https://orcid.org/" + orcId);
                                        run.setText(orcId);
                                        run.setUnderline(UnderlinePatterns.SINGLE);
                                        run.setColor("0000FF");
                                        paragraph.createRun().setText(")");
                                    }
                                    hasValue = true;
                                    hasMultiplicityItems = false;
                                } else {
                                    XWPFParagraph paragraph = addCellContent(extractValue, mainDocumentPart, field.getData().getFieldType().equals(FieldType.RICH_TEXT_AREA) ? ParagraphStyle.HTML : ParagraphStyle.TEXT, numId, indent, numOfRows, numOfCells, numOfValuesInCell);
                                    if (paragraph != null) {
                                        numOfValuesInCell++;
                                        if (orcidResearcher) {
                                            XWPFHyperlinkRun run = paragraph.createHyperlinkRun("https://orcid.org/" + orcId);
                                            run.setText(orcId);
                                            run.setUnderline(UnderlinePatterns.SINGLE);
                                            run.setColor("0000FF");
                                            paragraph.createRun().setText(")");
                                        }
                                        hasValue = true;
                                    }
                                }
                            }
                        }
                    }
                }
                numOfCells++;
            }
        }

        return hasValue;
    }

    private void createHypeLink(XWPFDocument mainDocumentPart, String format, String pidType, String pid, boolean hasMultiplicityItems, boolean isMultiAutoComplete) {
        PidLink pidLink = pidService.getPid(pidType);
        if (pidLink != null) {
            if (!hasMultiplicityItems) {
                XWPFParagraph paragraph = mainDocumentPart.createParagraph();
                paragraph.setIndentFromLeft(400 * indent);
                if (numId != null) {
                    paragraph.setNumID(numId);
                }
            }

            try {
                XWPFHyperlinkRun run = mainDocumentPart.getLastParagraph().createHyperlinkRun(pidLink.getLink().replace("{pid}", pid));

                if (isMultiAutoComplete) {
                    XWPFRun r = mainDocumentPart.getLastParagraph().createRun();
                    r.setText("• ");
                }

                run.setText(format);
                run.setUnderline(UnderlinePatterns.SINGLE);
                run.setColor("0000FF");
                run.setFontSize(11);
            } catch (Exception e) {
                String newFormat = (isMultiAutoComplete) ? "• " + format : format;
                if (hasMultiplicityItems) {
                    addParagraphContent(newFormat, mainDocumentPart, ParagraphStyle.TEXT, numId, indent);
                } else {
                    mainDocumentPart.getLastParagraph().createRun().setText(newFormat);
                }
            }
        } else {
            String newFormat = (isMultiAutoComplete) ? "• " + format : format;
            if (hasMultiplicityItems) {
                mainDocumentPart.getLastParagraph().createRun().setText(newFormat);
            } else {
                addParagraphContent(newFormat, mainDocumentPart, ParagraphStyle.TEXT, numId, indent);
            }
        }
    }

    private Boolean createFields(FieldSetModel fieldSetModel, PropertyDefinitionFieldSetItemModel propertyDefinitionFieldSetItemModel, XWPFDocument mainDocumentPart, Integer indent, Boolean createListing, boolean hasMultiplicityItems, VisibilityService visibilityService) {
        if (createListing) this.addListing(indent, false, false);
        boolean hasValue = false;
        List<FieldModel> tempFields = fieldSetModel.getFields().stream().sorted(Comparator.comparingInt(FieldModel::getOrdinal)).toList();
        for (FieldModel field : tempFields) {
            if (field.getIncludeInExport() && visibilityService.isVisible(field.getId(), propertyDefinitionFieldSetItemModel.getOrdinal())) {
                if (!createListing) {
                        org.opencdmp.commonmodels.models.description.FieldModel fieldValueModel = propertyDefinitionFieldSetItemModel.getFields().getOrDefault(field.getId(), null);
                        if (field.getData() != null) {
                            if (field.getData().getFieldType().equals(FieldType.UPLOAD)) {
                                boolean isImage = false;
                                for (UploadDataModel.UploadOptionModel type : ((UploadDataModel) field.getData()).getTypes()) {
                                    String fileFormat = type.getValue();
                                    if (IMAGE_TYPE_MAP.containsKey(fileFormat)) {
                                        isImage = true;
                                        break;
                                    }
                                }
                                if (isImage) {
                                    if (fieldValueModel.getTextValue() != null && !fieldValueModel.getTextValue().isEmpty()) {
                                        XWPFParagraph paragraph = addParagraphContent(fieldValueModel.getFile(), mainDocumentPart, ParagraphStyle.IMAGE, numId, 0); //TODO
                                        if (paragraph != null) {
                                            hasValue = true;
                                        }
                                        if (hasMultiplicityItems) {
                                            hasMultiplicityItems = false;
                                        }
                                    }
                                }
                            } else if (fieldValueModel != null) {
                                this.indent = indent;
                                boolean isMultiAutoComplete = false;
                                boolean isResearcher = false;
                                boolean isOrganization = false;
                                boolean isExternalDataset = false;
                                boolean isPublication = false;
                                if (field.getData() instanceof LabelAndMultiplicityDataModel) {
                                    isMultiAutoComplete = ((LabelAndMultiplicityDataModel) field.getData()).getMultipleSelect() != null && ((LabelAndMultiplicityDataModel) field.getData()).getMultipleSelect();
                                }
                                if (field.getData() instanceof SelectDataModel) {
                                    isMultiAutoComplete = ((SelectDataModel) field.getData()).getMultipleSelect() != null && ((SelectDataModel) field.getData()).getMultipleSelect();
                                }
                                if (field.getData() instanceof ReferenceTypeDataModel) {
                                    isMultiAutoComplete = ((ReferenceTypeDataModel) field.getData()).getMultipleSelect() != null && ((ReferenceTypeDataModel) field.getData()).getMultipleSelect();
                                    isResearcher = ((ReferenceTypeDataModel) field.getData()).getReferenceType().getCode().equals(this.wordFileTransformerServiceProperties.getResearcherReferenceCode());
                                    isOrganization = ((ReferenceTypeDataModel) field.getData()).getReferenceType().getCode().equals(this.wordFileTransformerServiceProperties.getOrganizationReferenceCode());
                                    isExternalDataset = ((ReferenceTypeDataModel) field.getData()).getReferenceType().getCode().equals(this.wordFileTransformerServiceProperties.getDatasetReferenceCode());
                                    isPublication = ((ReferenceTypeDataModel) field.getData()).getReferenceType().getCode().equals(this.wordFileTransformerServiceProperties.getPublicationReferenceCode());
                                }

                                if (isOrganization || isExternalDataset || isPublication) {
                                    if (fieldValueModel.getReferences() != null) {
                                        for (ReferenceModel referenceModel : fieldValueModel.getReferences()) {
                                            String label = "";
                                            if (referenceModel.getLabel() != null && !referenceModel.getLabel().isBlank()) {
                                                label =  referenceModel.getLabel();
                                            } else if (referenceModel.getDescription() != null && !referenceModel.getDescription().isBlank()) {
                                                label = (label.isBlank() ? "" : " ") + referenceModel.getDescription();
                                            }
                                            ReferenceFieldModel fieldModel = referenceModel.getDefinition() != null && referenceModel.getDefinition().getFields() != null &&  !referenceModel.getDefinition().getFields().isEmpty() ? referenceModel.getDefinition().getFields().stream().filter(x -> x.getCode().equals("pidTypeField")).findFirst().orElse(null) : null;
                                            createHypeLink(mainDocumentPart, label, fieldModel != null ? fieldModel.getValue() : null, referenceModel.getReference(), hasMultiplicityItems, isMultiAutoComplete && fieldValueModel.getReferences().size() > 1);
                                        }
                                        if (hasMultiplicityItems) hasMultiplicityItems = false;

                                        hasValue = true;
                                    }

                                } else {
                                    List<String> extractValues = this.extractValues(field, fieldValueModel);

                                    if (!extractValues.isEmpty()){
                                        for (String extractValue : extractValues){
                                            boolean orcidResearcher = false;
                                            String orcId = null;
                                            if (isResearcher && extractValue.contains("orcid:")) {
                                                orcId = extractValue.substring(extractValue.indexOf(':') + 1, extractValue.indexOf(')'));
                                                extractValue = extractValue.substring(0, extractValue.indexOf(':') + 1) + " ";
                                                orcidResearcher = true;
                                            }
                                            if (extractValues.size() > 1) extractValue = "• " + extractValue;
                                            if (hasMultiplicityItems) {
                                                mainDocumentPart.getLastParagraph().createRun().setText(extractValue);
                                                if (orcidResearcher) {
                                                    XWPFHyperlinkRun run = mainDocumentPart.getLastParagraph().createHyperlinkRun("https://orcid.org/" + orcId);
                                                    run.setText(orcId);
                                                    run.setUnderline(UnderlinePatterns.SINGLE);
                                                    run.setColor("0000FF");
                                                    mainDocumentPart.getLastParagraph().createRun().setText(")");
                                                }
                                                hasValue = true;
                                                hasMultiplicityItems = false;
                                            } else {
                                                XWPFParagraph paragraph = addParagraphContent(extractValue, mainDocumentPart, field.getData().getFieldType().equals(FieldType.RICH_TEXT_AREA) ? ParagraphStyle.HTML : ParagraphStyle.TEXT, numId, indent);
                                                if (paragraph != null) {
                                                    if (orcidResearcher) {
                                                        XWPFHyperlinkRun run = paragraph.createHyperlinkRun("https://orcid.org/" + orcId);
                                                        run.setText(orcId);
                                                        run.setUnderline(UnderlinePatterns.SINGLE);
                                                        run.setColor("0000FF");
                                                        paragraph.createRun().setText(")");
                                                    }
                                                    hasValue = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                }
            }
        }
        return hasValue;
    }

    private XWPFParagraph addCellContent(Object content, XWPFTableRow mainDocumentPart, ParagraphStyle style, BigInteger numId, int indent, int numOfRows, int numOfCells, int numOfValuesInCell) {
        if (content == null) return null;
        if (content instanceof String && ((String) content).isEmpty())  return null;
        
        this.indent = indent;
        XWPFTableCell cell;
        if (numOfRows > 0 || numOfValuesInCell > 0) {
            cell = mainDocumentPart.getCell(numOfCells);
        } else {
            cell = mainDocumentPart.createCell();
        }
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf("CENTER"));
        if (numOfValuesInCell == 0) {
            cell.removeParagraph(0);
        }

        XWPFParagraph paragraph = this.optionsInTable.get(style).apply(cell, content);
        if (paragraph != null) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            paragraph.setSpacingBefore(100);
            if (numId != null) {
                paragraph.setNumID(numId);
            }
            return paragraph;
        }
        return null;
    }

    @Override
    public XWPFParagraph addParagraphContent(Object content, XWPFDocument mainDocumentPart, ParagraphStyle style, BigInteger numId, int indent) {
        if (content != null) {
            if (content instanceof String && ((String)content).isEmpty()) {
                return null;
            }
            this.indent = indent;
            XWPFParagraph paragraph = this.options.get(style).apply(mainDocumentPart, content);
            if (paragraph != null) {
                paragraph.setIndentFromLeft(400*indent);
                if (numId != null) {
                    paragraph.setNumID(numId);
                }
                return paragraph;
            }
        }
        return null;
    }

    private void addListing(int indent, boolean question, Boolean hasIndication) {
        CTLvl cTLvl = this.cTAbstractNum.addNewLvl();

        if (question) {
            cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
            cTLvl.setIlvl(BigInteger.valueOf(indent));
        } else {
            if (hasIndication) {
                cTLvl.addNewNumFmt().setVal(STNumberFormat.DECIMAL);
                cTLvl.setIlvl(BigInteger.valueOf(indent));
            } else {
                cTLvl.addNewNumFmt().setVal(STNumberFormat.NONE);
                cTLvl.setIlvl(BigInteger.valueOf(indent));
            }
        }
    }

    private List<String> extractValues(FieldModel field, org.opencdmp.commonmodels.models.description.FieldModel fieldValueModel) {
        List<String> values = new ArrayList<>();
        if (fieldValueModel == null || field == null || field.getData() == null) {
            return values;
        }
        switch (field.getData().getFieldType()) {
            case REFERENCE_TYPES: {
                if (fieldValueModel.getReferences() != null && !fieldValueModel.getReferences().isEmpty()) {
                    for (ReferenceModel referenceModel : fieldValueModel.getReferences()) {
                        if (referenceModel != null) {
                            String label = "";
                            if (referenceModel.getLabel() != null && !referenceModel.getLabel().isBlank()) {
                                label =  referenceModel.getLabel();
                            } else if (referenceModel.getDescription() != null && !referenceModel.getDescription().isBlank()) {
                                label = (label.isBlank() ? "" : " ") + referenceModel.getDescription();
                            }
                            if (!label.isBlank()) values.add(label);
                        }
                    }
                }
                break;
            }
            case TAGS:
                if (fieldValueModel.getTextListValue() != null && !fieldValueModel.getTextListValue().isEmpty()) {
                    values.addAll(fieldValueModel.getTextListValue());
                }
                break;
            case SELECT: {
                if (fieldValueModel.getTextListValue() != null && !fieldValueModel.getTextListValue().isEmpty()) {
                    SelectDataModel selectDataModel = (SelectDataModel) field.getData();
                    if (selectDataModel != null && selectDataModel.getOptions() != null && !selectDataModel.getOptions().isEmpty()) {
                        for (SelectDataModel.OptionModel option : selectDataModel.getOptions()) {
                            if (fieldValueModel.getTextListValue().contains(option.getValue()) || fieldValueModel.getTextListValue().contains(option.getLabel())) values.add(option.getLabel());
                        }
                    }
                }
                break;
            }
            case BOOLEAN_DECISION:
                if (fieldValueModel.getBooleanValue() != null && fieldValueModel.getBooleanValue()) values.add("Yes");
                if (fieldValueModel.getBooleanValue() != null && !fieldValueModel.getBooleanValue()) values.add("No");
                break;
            case RADIO_BOX:
                RadioBoxDataModel radioBoxDataModel = (RadioBoxDataModel) field.getData();
                if (fieldValueModel.getTextValue() != null && radioBoxDataModel != null && radioBoxDataModel.getOptions() != null) {
                    for (RadioBoxDataModel.RadioBoxOptionModel option : radioBoxDataModel.getOptions()) {
                        if (option.getValue().equals(fieldValueModel.getTextValue()) || option.getLabel().equals(fieldValueModel.getTextValue())) {
                            values.add(option.getLabel());
                            break;
                        }
                    }
                }
                break;
            case CHECK_BOX: {
                LabelDataModel checkBoxData = (LabelDataModel) field.getData();
                if (fieldValueModel.getBooleanValue() != null && fieldValueModel.getBooleanValue() && checkBoxData != null && checkBoxData.getLabel() != null) values.add(checkBoxData.getLabel());
                break;
            }
            case DATE_PICKER: {
                if (fieldValueModel.getDateValue() != null) values.add(DateTimeFormatter.ofPattern("yyyy-MM-dd").withZone(ZoneId.systemDefault()).format(fieldValueModel.getDateValue()));
                break;
            }
            case FREE_TEXT:
            case TEXT_AREA:
            case RICH_TEXT_AREA: {
                if (fieldValueModel.getTextValue() != null && !fieldValueModel.getTextValue().isBlank()) values.add(fieldValueModel.getTextValue());
                break;
            }
            case DATASET_IDENTIFIER:
            case VALIDATION: {
                if (fieldValueModel.getExternalIdentifier() != null) {
                    values.add("id: " + fieldValueModel.getExternalIdentifier().getIdentifier() + ", Type: " + fieldValueModel.getExternalIdentifier().getType());
                }
                break;
            }
            case UPLOAD:
            case INTERNAL_ENTRIES_DESCRIPTIONS:
            case INTERNAL_ENTRIES_PlANS:
                break;
            default:
                throw new MyApplicationException("Invalid type " + field.getData().getFieldType());
        }
        
        return values;
    }

    @Override
    public int findPosOfPoweredBy(XWPFDocument document) {
        if (document == null) throw new MyApplicationException("Document required");
        if (document.getParagraphs() == null) return -1;

        for (XWPFParagraph p : document.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null) {
                        if (text.equals("Powered by")) {
                            return document.getPosOfParagraph(p) - 1;
                        }
                    }
                }
            }
        }
        return -1;
    }

    private List<ReferenceModel> getReferenceModelOfTypeCode(PlanModel plan, String code) {
        List<ReferenceModel> response = new ArrayList<>();
        if (plan.getReferences() == null) return response;
        for (PlanReferenceModel planReferenceModel : plan.getReferences()) {
            if (planReferenceModel.getReference() != null && planReferenceModel.getReference().getType() != null && planReferenceModel.getReference().getType().getCode() != null && planReferenceModel.getReference().getType().getCode().equals(code)) {
                response.add(planReferenceModel.getReference());
            }
        }
        return response;
    }

    @Override
    public void fillFirstPage(PlanModel planEntity, DescriptionModel descriptionModel, XWPFDocument document, boolean isDescription) {
        if (planEntity == null) throw new MyApplicationException("planEntity required");
        if (document == null) throw new MyApplicationException("Document required");

        int parPos = 0;
        int descrParPos = -1;
        List<ReferenceModel> grants = this.getReferenceModelOfTypeCode(planEntity, this.wordFileTransformerServiceProperties.getGrantReferenceCode());
        List<ReferenceModel> researchers = this.getReferenceModelOfTypeCode(planEntity, this.wordFileTransformerServiceProperties.getResearcherReferenceCode());
        List<ReferenceModel> organizations = this.getReferenceModelOfTypeCode(planEntity, this.wordFileTransformerServiceProperties.getOrganizationReferenceCode());
        List<ReferenceModel> funders = this.getReferenceModelOfTypeCode(planEntity, this.wordFileTransformerServiceProperties.getFunderReferenceCode());

        XWPFParagraph descrPar = null;
        for (XWPFParagraph p : document.getParagraphs()) {

            this.replaceTextSegment(p, "'{ARGOS.DMP.TITLE}'", planEntity.getLabel());
            this.replaceTextSegment(p, "'{ARGOS.DMP.VERSION}'", "Version " + planEntity.getVersion());
            if (descriptionModel != null) {
                this.replaceTextSegment(p, "'{ARGOS.DATASET.TITLE}'", descriptionModel.getLabel());
            }

            StringBuilder researchersNames = new StringBuilder();
            int i = 0;
            for (ReferenceModel researcher : researchers) {
                i++;
                researchersNames.append(researcher.getLabel()).append(i < researchers.size() ? ", " : "");
            }
            this.replaceTextSegment(p, "'{ARGOS.DMP.RESEARCHERS}'", researchersNames.toString(), 15);

            StringBuilder organisationsNames = new StringBuilder();
            i = 0;
            for (ReferenceModel organisation : organizations) {
                i++;
                organisationsNames.append(organisation.getLabel()).append(i < organizations.size() ? ", " : "");
            }
            this.replaceTextSegment(p, "'{ARGOS.DMP.ORGANIZATIONS}'", organisationsNames.toString(), 15);

            if (this.textSegmentExists(p, "'{ARGOS.DMP.DESCRIPTION}'")) {
                descrParPos = parPos;
                descrPar = p;
                this.replaceTextSegment(p, "'{ARGOS.DMP.DESCRIPTION}'", "");
            }
            if (this.textSegmentExists(p, "'{ARGOS.DATASET.DESCRIPTION}'")) {
                descrParPos = parPos;
                descrPar = p;
                this.replaceTextSegment(p, "'{ARGOS.DATASET.DESCRIPTION}'", "");
            }
        }
        if ((descrParPos != -1) &&  (planEntity.getDescription() != null) && !isDescription) {
            XmlCursor cursor = descrPar.getCTP().newCursor();
            cursor.toNextSibling();
            Document htmlDoc = Jsoup.parse((planEntity.getDescription()).replaceAll("\n", "<br>"));
            HtmlToWorldBuilder htmlToWorldBuilder = new HtmlToWorldBuilder(descrPar, 0, cursor);
            NodeTraversor.traverse(htmlToWorldBuilder, htmlDoc);
        }
        if ((descrParPos != -1) && (descriptionModel != null) && (descriptionModel.getDescription() != null) && isDescription) {
            XmlCursor cursor = descrPar.getCTP().newCursor();
            cursor.toNextSibling();
            Document htmlDoc = Jsoup.parse((descriptionModel.getDescription()).replaceAll("\n", "<br>"));
            HtmlToWorldBuilder htmlToWorldBuilder = new HtmlToWorldBuilder(descrPar, 0, cursor);
            NodeTraversor.traverse(htmlToWorldBuilder, htmlDoc);
        }


        XWPFTable tbl = document.getTables().getFirst();
        Iterator<XWPFTableRow> it = tbl.getRows().iterator();
        it.next(); // skip first row
        if (it.hasNext() && !funders.isEmpty()) {
            XWPFParagraph p = it.next().getCell(0).getParagraphs().getFirst();
            XWPFRun run = p.createRun();
            run.setText(funders.getFirst().getLabel());
            run.setFontSize(15);
            p.setAlignment(ParagraphAlignment.CENTER);
        }
        it = tbl.getRows().iterator();
        it.next();
        if (it.hasNext() && !grants.isEmpty()) {
            XWPFParagraph p = it.next().getCell(1).getParagraphs().getFirst();
            XWPFRun run = p.createRun();
            String text = grants.getFirst().getLabel();
            String reference = grants.getFirst().getReference();
            if (reference != null) {
                String[] parts = reference.split("::");
                text += parts.length > 1 ? "/ No " + parts[parts.length - 1] : "";
            }
            run.setText(text);
            run.setFontSize(15);
            p.setAlignment(ParagraphAlignment.CENTER);
        }
    }

    private boolean textSegmentExists(XWPFParagraph paragraph, String textToFind) {
        PositionInParagraph startPos = new PositionInParagraph(0, 0, 0);
        return this.searchText(paragraph, textToFind, startPos) != null;
    }

    private void replaceTextSegment(XWPFParagraph paragraph, String textToFind, String replacement) {
        this.replaceTextSegment(paragraph, textToFind, replacement, null);
    }
    
    private void replaceTextSegment(XWPFParagraph paragraph, String textToFind, String replacement, Integer fontSize) {
        TextSegment foundTextSegment;
        PositionInParagraph startPos = new PositionInParagraph(0, 0, 0);
        while((foundTextSegment = this.searchText(paragraph, textToFind, startPos)) != null) { // search all text segments having text to find

            logger.debug(foundTextSegment.getBeginRun()+":"+foundTextSegment.getBeginText()+":"+foundTextSegment.getBeginChar());
            logger.debug(foundTextSegment.getEndRun()+":"+foundTextSegment.getEndText()+":"+foundTextSegment.getEndChar());

            // maybe there is text before textToFind in begin run
            XWPFRun beginRun = paragraph.getRuns().get(foundTextSegment.getBeginRun());
            String textInBeginRun = beginRun.getText(foundTextSegment.getBeginText());
            String textBefore = textInBeginRun.substring(0, foundTextSegment.getBeginChar()); // we only need the text before

            // maybe there is text after textToFind in end run
            XWPFRun endRun = paragraph.getRuns().get(foundTextSegment.getEndRun());
            String textInEndRun = endRun.getText(foundTextSegment.getEndText());
            String textAfter = textInEndRun.substring(foundTextSegment.getEndChar() + 1); // we only need the text after

            if (foundTextSegment.getEndRun() == foundTextSegment.getBeginRun()) {
                textInBeginRun = textBefore + replacement + textAfter; // if we have only one run, we need the text before, then the replacement, then the text after in that run
            } else {
                textInBeginRun = textBefore + replacement; // else we need the text before followed by the replacement in begin run
                endRun.setText(textAfter, foundTextSegment.getEndText()); // and the text after in end run
            }

            beginRun.setText(textInBeginRun, foundTextSegment.getBeginText());
            if (fontSize != null) {
                beginRun.setFontSize(fontSize);
            }
            // runs between begin run and end run needs to be removed
            for (int runBetween = foundTextSegment.getEndRun() - 1; runBetween > foundTextSegment.getBeginRun(); runBetween--) {
                paragraph.removeRun(runBetween); // remove not needed runs
            }

        }
    }

    private TextSegment searchText(XWPFParagraph paragraph, String searched, PositionInParagraph startPos) {
        int startRun = startPos.getRun(),
                startText = startPos.getText(),
                startChar = startPos.getChar();
        int beginRunPos = 0, candCharPos = 0;
        boolean newList = false;

        //CTR[] rArray = paragraph.getRArray(); //This does not contain all runs. It lacks hyperlink runs for ex.
        java.util.List<XWPFRun> runs = paragraph.getRuns();

        int beginTextPos = 0, beginCharPos = 0; //must be outside the for loop

        //for (int runPos = startRun; runPos < rArray.length; runPos++) {
        for (int runPos = startRun; runPos < runs.size(); runPos++) {
            //int beginTextPos = 0, beginCharPos = 0, textPos = 0, charPos; //int beginTextPos = 0, beginCharPos = 0 must be outside the for loop
            int textPos = 0, charPos;
            //CTR ctRun = rArray[runPos];
            CTR ctRun = runs.get(runPos).getCTR();
            XmlCursor c = ctRun.newCursor();
	        try (c) {
		        c.selectPath("./*");
		        while (c.toNextSelection()) {
			        XmlObject o = c.getObject();
			        if (o instanceof CTText) {
				        if (textPos >= startText) {
					        String candidate = ((CTText) o).getStringValue();
					        if (runPos == startRun) {
						        charPos = startChar;
					        } else {
						        charPos = 0;
					        }

					        for (; charPos < candidate.length(); charPos++) {
						        if ((candidate.charAt(charPos) == searched.charAt(0)) && (candCharPos == 0)) {
							        beginTextPos = textPos;
							        beginCharPos = charPos;
							        beginRunPos = runPos;
							        newList = true;
						        }
						        if (candidate.charAt(charPos) == searched.charAt(candCharPos)) {
							        if (candCharPos + 1 < searched.length()) {
								        candCharPos++;
							        } else if (newList) {
								        TextSegment segment = new TextSegment();
								        segment.setBeginRun(beginRunPos);
								        segment.setBeginText(beginTextPos);
								        segment.setBeginChar(beginCharPos);
								        segment.setEndRun(runPos);
								        segment.setEndText(textPos);
								        segment.setEndChar(charPos);
								        return segment;
							        }
						        } else {
							        candCharPos = 0;
						        }
					        }
				        }
				        textPos++;
			        } else if (o instanceof CTProofErr) {
				        c.removeXml();
			        } else if (o instanceof CTRPr) {
				        //do nothing
			        } else {
				        candCharPos = 0;
			        }
		        }
	        }
        }
        return null;
    }

    @Override
    public void fillFooter(PlanModel planEntity, DescriptionModel descriptionModel, XWPFDocument document) {
        if (planEntity == null) throw new MyApplicationException("planEntity required");

        List<ReferenceModel> licences = this.getReferenceModelOfTypeCode(planEntity, this.wordFileTransformerServiceProperties.getLicenceReferenceCode());
        document.getFooterList().forEach(xwpfFooter -> {
            for (XWPFParagraph p : xwpfFooter.getParagraphs()) {
                if (p != null) {
                    this.replaceTextSegment(p, "'{ARGOS.DMP.TITLE}'", planEntity.getLabel());
                    if (descriptionModel != null) {
                        this.replaceTextSegment(p, "'{ARGOS.DATASET.TITLE}'", descriptionModel.getLabel());
                    }
                    if (!licences.isEmpty() && licences.getFirst().getReference() != null && !licences.getFirst().getReference().isBlank()) {
                        this.replaceTextSegment(p, "'{ARGOS.DMP.LICENSE}'", licences.getFirst().getReference());
                    } else {
                        this.replaceTextSegment(p, "'{ARGOS.DMP.LICENSE}'", "License: -");
                    }
                    if (planEntity.getEntityDois() != null && !planEntity.getEntityDois().isEmpty()) {
                        this.replaceTextSegment(p, "'{ARGOS.DMP.DOI}'", planEntity.getEntityDois().getFirst().getDoi());
                    } else {
                        this.replaceTextSegment(p, "'{ARGOS.DMP.DOI}'", "-");
                    }
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy").withZone(ZoneId.systemDefault());
                    this.replaceTextSegment(p, "'{ARGOS.DMP.LAST_MODIFIED}'", formatter.format(planEntity.getUpdatedAt()));
                }
            }
        });
    }
}
