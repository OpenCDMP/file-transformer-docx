package org.opencdmp.filetransformer.docx.service.template;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.BlockTemplate;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.opencdmp.commonmodels.enums.PluginEntityType;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationIssueLevel;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationIssueModel;

import java.io.ByteArrayInputStream;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/**
 * Checks a template: that it does not use the unsupported '{OPENCDMP.X}' placeholders, that poi-tl can parse it, that its top-level tags use the template
 * vocabulary, and that it renders the questionnaire content at all. Tags inside blocks are relative
 * to the block's item and are only checked when they start with a global root.
 */
public final class TemplateValidator {

    private static final Set<String> ROOTS = Set.of("plan", "asset", "chapters", "assets", "layout", "labels", "export", "branding", "toc", "pageBreak");
    private static final RenderPolicy NO_OP = (template, data, xwpfTemplate) -> {
    };

    private TemplateValidator() {
    }

    public static void validate(byte[] template, PluginEntityType entityType, List<TemplateValidationIssueModel> issues) {
        List<MetaTemplate> tags;
        try (XWPFDocument document = new XWPFDocument(new ByteArrayInputStream(template))) {
            if (TemplateTexts.hasOldPlaceholders(document)) {
                issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Error, "unsupported-placeholders",
                        "The template uses '{OPENCDMP.X}' placeholders, which are no longer supported. Exports will use the built-in template instead. "
                                + "Start from the built-in template and use {{...}} tags."));
                return;
            }
            Configure configure = Configure.builder().addPlugin(TemplateRenderer.CONTENT_SIGN, NO_OP).build();
            tags = XWPFTemplate.compile(document, configure).getElementTemplates();
        } catch (Exception e) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Error, "unparseable-template",
                    "The {{...}} tags of the template could not be read: " + e.getMessage()));
            return;
        }

        Set<String> used = new LinkedHashSet<>();
        Set<String> topLevel = new LinkedHashSet<>();
        Set<String> unknown = new LinkedHashSet<>();
        collect(tags, true, used, topLevel, unknown);
        for (String tag : unknown) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Warning, "unknown-tag",
                    "{{" + tag + "}} does not start with a known name (" + String.join(", ", ROOTS) + ") and will render empty."));
        }

        boolean plan = entityType != PluginEntityType.Description;
        if (plan && topLevel.stream().anyMatch(x -> x.equals("asset") || x.startsWith("asset."))) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Warning, "entity-mismatch",
                    "'asset' is only set in asset exports; in a plan template loop over 'chapters' or 'assets' instead."));
        }
        if (!plan && topLevel.stream().anyMatch(x -> x.equals("chapters"))) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Warning, "entity-mismatch",
                    "'chapters' is only set in plan exports; in an asset template render {{>asset.body}}."));
        }
        boolean rendersBody = plan ? used.contains("chapters") || used.contains("assets") : used.contains("asset.body");
        if (!rendersBody) {
            issues.add(new TemplateValidationIssueModel(TemplateValidationIssueLevel.Warning, "missing-body",
                    plan ? "The template never loops over {{?chapters}}, so exports will not include the plan's sections and assets."
                            : "The template never renders {{>asset.body}}, so exports will not include the asset's answers."));
        }
    }

    private static void collect(List<MetaTemplate> templates, boolean topLevel, Set<String> used, Set<String> topLevelNames, Set<String> unknown) {
        for (MetaTemplate template : templates) {
            String name = null;
            if (template instanceof BlockTemplate block) {
                name = block.getStartMark() != null ? block.getStartMark().getTagName() : null;
            } else if (template instanceof ElementTemplate element) {
                name = element.getTagName();
            }
            if (name != null && !name.isBlank()) {
                String trimmed = name.trim();
                String root = trimmed.split("\\.", 2)[0];
                if (topLevel) topLevelNames.add(trimmed);
                if (topLevel || ROOTS.contains(root)) {
                    used.add(trimmed);
                    if (!ROOTS.contains(root) && !root.startsWith("_")) unknown.add(trimmed);
                }
            }
            if (template instanceof BlockTemplate block && block.getTemplates() != null) {
                collect(new ArrayList<>(block.getTemplates()), false, used, topLevelNames, unknown);
            }
        }
    }
}
