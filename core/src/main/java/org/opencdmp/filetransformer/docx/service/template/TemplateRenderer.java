package org.opencdmp.filetransformer.docx.service.template;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.render.compute.DefaultELRenderDataCompute;
import com.deepoove.poi.render.compute.EnvModel;
import com.deepoove.poi.render.compute.RenderDataCompute;
import gr.cite.tools.exception.MyApplicationException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.render.ContentRenderPolicy;
import org.opencdmp.filetransformer.docx.service.template.render.ContentRenderPolicy.TocEntry;
import org.opencdmp.filetransformer.docx.service.template.render.TemplateStyles;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordExportOptions;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Renders '{{...}}' templates with poi-tl against an {@link ExportView}. Expressions use poi-tl's
 * logic-less default expression language only (never SpEL): templates are uploaded by tenant
 * administrators and must not be able to run code.
 */
@Component
public class TemplateRenderer {
    private static final Logger logger = LoggerFactory.getLogger(TemplateRenderer.class);

    /** The sign of the content tag: {@code {{>plan.description}}}. */
    public static final char CONTENT_SIGN = '>';
    /** Root names that resolve from inside any loop or block. */
    static final Set<String> GLOBAL_ROOTS = Set.of("plan", "layout", "labels", "export", "branding", "toc", "pageBreak");

    private static final String BUILT_IN_PLAN_TEMPLATE = "classpath:documents/plan.docx";
    private static final String BUILT_IN_ASSET_TEMPLATE = "classpath:documents/description.docx";

    private final ExportViewBuilder viewBuilder;
    private final ResourceLoader resourceLoader;
    private final WordFileTransformerServiceProperties properties;
    private final ConcurrentHashMap<String, byte[]> builtIns = new ConcurrentHashMap<>();

    public TemplateRenderer(ExportViewBuilder viewBuilder, ResourceLoader resourceLoader, WordFileTransformerServiceProperties properties) {
        this.viewBuilder = viewBuilder;
        this.resourceLoader = resourceLoader;
        this.properties = properties;
    }

    public XWPFDocument renderPlan(XWPFDocument template, PlanModel plan, WordExportOptions options) {
        ExportView view = this.viewBuilder.forPlan(plan, options);
        List<TocEntry> toc = new ArrayList<>();
        for (ExportView.ChapterView chapter : view.chapters) {
            if (chapter.section != null) toc.add(new TocEntry(1, chapter.section.heading));
            if (chapter.group != null) toc.add(new TocEntry(1, chapter.group.heading != null ? chapter.group.heading : chapter.group.title));
            if (chapter.asset != null) toc.add(new TocEntry(2, chapter.asset.heading));
        }
        return this.render(template, view, options, toc, this.orBuiltIn(this.properties.getWordPlanTemplate(), BUILT_IN_PLAN_TEMPLATE));
    }

    public XWPFDocument renderAsset(XWPFDocument template, DescriptionModel asset, WordExportOptions options) {
        ExportView view = this.viewBuilder.forAsset(asset, options);
        List<TocEntry> toc = new ArrayList<>();
        if (view.asset.body != null) {
            for (Body.Node node : view.asset.body.nodes()) {
                if (node instanceof Body.SectionNode section) toc.add(new TocEntry(1, section.number() != null ? section.number() + " " + section.title() : section.title()));
            }
        }
        return this.render(template, view, options, toc, this.orBuiltIn(this.properties.getWordDescriptionTemplate(), BUILT_IN_ASSET_TEMPLATE));
    }

    private String orBuiltIn(String configured, String builtIn) {
        return configured != null && !configured.isBlank() ? configured : builtIn;
    }

    private XWPFDocument render(XWPFDocument template, ExportView view, WordExportOptions options, List<TocEntry> toc, String builtIn) {
        try {
            // missing contract styles come from the default template; serialize so poi-tl sees them
            XWPFDocument reference = this.builtIn(builtIn);
            TemplateStyles.ensure(template, reference);
            ByteArrayOutputStream prepared = new ByteArrayOutputStream();
            template.write(prepared);

            Configure configure = Configure.builder()
                    .addPlugin(CONTENT_SIGN, new ContentRenderPolicy(view.labels, options, toc))
                    .setRenderDataComputeFactory(model -> new GlobalFallbackCompute(model, view))
                    .build();
            XWPFTemplate rendered = XWPFTemplate.compile(new ByteArrayInputStream(prepared.toByteArray()), configure).render(view);
            return rendered.getXWPFDocument();
        } catch (Exception e) {
            logger.error("could not render the DOCX template", e);
            throw new MyApplicationException("could not render the DOCX template: " + e.getMessage());
        }
    }

    private XWPFDocument builtIn(String location) {
        try {
            byte[] bytes = this.builtIns.computeIfAbsent(location, key -> {
                Resource resource = this.resourceLoader.getResource(key);
                try (InputStream input = resource.getInputStream()) {
                    return input.readAllBytes();
                } catch (Exception e) {
                    throw new IllegalStateException(e);
                }
            });
            return new XWPFDocument(new ByteArrayInputStream(bytes));
        } catch (Exception e) {
            logger.warn("default template {} unavailable, styles missing from custom templates will not be added", location, e);
            return null;
        }
    }

    /** Resolves an expression in the current scope, falling back to the root view for {@link #GLOBAL_ROOTS}. */
    private static final class GlobalFallbackCompute implements RenderDataCompute {
        private final EnvModel model;
        private final Object root;
        private final RenderDataCompute local;

        GlobalFallbackCompute(EnvModel model, Object root) {
            this.model = model;
            this.root = root;
            this.local = new DefaultELRenderDataCompute(model, false);
        }

        @Override
        public Object compute(String expression) {
            Object value = this.local.compute(expression);
            if (value != null || this.model.getRoot() == this.root || expression == null) return value;
            int dot = expression.indexOf('.');
            String first = (dot < 0 ? expression : expression.substring(0, dot)).trim();
            if (!GLOBAL_ROOTS.contains(first)) return null;
            return new DefaultELRenderDataCompute(EnvModel.ofModel(this.root), false).compute(expression);
        }
    }
}
