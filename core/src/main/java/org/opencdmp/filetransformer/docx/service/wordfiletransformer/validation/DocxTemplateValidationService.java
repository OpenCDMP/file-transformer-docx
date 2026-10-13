package org.opencdmp.filetransformer.docx.service.wordfiletransformer.validation;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.opencdmp.commonmodels.enums.PluginEntityType;
import org.opencdmp.filetransformer.docx.service.template.TemplateValidator;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationIssueLevel;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationIssueModel;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationResultModel;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.util.List;

/**
 * Validates an uploaded DOCX template: that it can be read as a DOCX document, then its
 * '{{...}}' tags through {@link TemplateValidator}.
 */
@Component
public class DocxTemplateValidationService {

    public TemplateValidationResultModel validate(byte[] template, PluginEntityType entityType) {
        TemplateValidationResultModel result = new TemplateValidationResultModel();
        List<TemplateValidationIssueModel> issues = result.getIssues();

        try (XWPFDocument ignored = new XWPFDocument(new ByteArrayInputStream(template))) {
            // readable
        } catch (Exception e) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Error, "unreadable",
                    "The file could not be read as a DOCX document."));
            result.setValid(false);
            return result;
        }

        TemplateValidator.validate(template, entityType, issues);
        result.setValid(issues.stream().noneMatch(x -> TemplateValidationIssueLevel.Error.equals(x.getLevel())));
        return result;
    }
}
