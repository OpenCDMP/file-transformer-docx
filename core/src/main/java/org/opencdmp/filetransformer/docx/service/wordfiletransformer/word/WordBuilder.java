package org.opencdmp.filetransformer.docx.service.wordfiletransformer.word;

import org.apache.xmlbeans.XmlCursor;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.description.PropertyDefinitionModel;
import org.opencdmp.commonmodels.models.descriptiotemplate.DescriptionTemplateModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.reference.ReferenceModel;
import org.opencdmp.filetransformer.docx.model.enums.ParagraphStyle;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.visibility.VisibilityService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.IOException;
import java.math.BigInteger;
import java.util.List;
import java.util.UUID;

public interface WordBuilder {
	void build(XWPFDocument document, DescriptionTemplateModel descriptionTemplate, PropertyDefinitionModel propertyDefinitionModel, VisibilityService visibilityService, XWPFParagraph paragraphCursor) throws IOException;

	XWPFParagraph addParagraphContent(Object content, XWPFDocument mainDocumentPart, ParagraphStyle style, BigInteger numId, int indent, XmlCursor cursor, boolean isReference,  boolean orcidResearcher, org.opencdmp.commonmodels.models.description.FieldModel fieldValueModel, UUID referenceId);

	XWPFParagraph findParagraphFormCode(XWPFDocument document, String code);

	void fillFirstPage(PlanModel planEntity, DescriptionModel descriptionModel, XWPFDocument document, boolean isDescription);

	void fillFooter(PlanModel planEntity, DescriptionModel descriptionModel, XWPFDocument document);

	void fillHeader(PlanModel planEntity, DescriptionModel descriptionModel, XWPFDocument document);

	void addReferenceFieldDefinitionToParagraph(List<ReferenceModel> references, UUID referenceId, XWPFParagraph paragraph);

}
