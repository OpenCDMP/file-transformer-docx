package org.opencdmp.filetransformer.docx.service.template.view;

import com.deepoove.poi.data.PictureRenderData;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.content.Facts;
import org.opencdmp.filetransformer.docx.service.template.content.PageBreak;
import org.opencdmp.filetransformer.docx.service.template.content.RichText;
import org.opencdmp.filetransformer.docx.service.template.content.TableOfContents;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * The data a '{{...}}' template is rendered against: the template vocabulary.
 * <p>
 * poi-tl reads <b>fields</b> (not getters), so every name below is exactly what a template author
 * writes: {@code {{plan.title}}}, {@code {{?chapters}}...{{/chapters}}}, {@code {{>asset.body}}}.
 * Values are fully resolved (labels instead of ids, formatted dates, visibility applied, unanswered
 * questions pruned) so templates never need logic. Fields typed as content ({@link RichText},
 * {@link Body}, {@link Facts}, ...) are rendered with the {@code >} tag, e.g. {@code {{>plan.description}}}.
 * <p>
 * The root names {@code plan}, {@code layout}, {@code labels}, {@code export}, {@code branding},
 * {@code toc} and {@code pageBreak} resolve from anywhere, including inside loops.
 */
public class ExportView {

    /** Always present; for an asset export this is the asset's plan. */
    public PlanView plan;
    /** Asset export only: the exported asset. */
    public AssetView asset;
    /** Plan export only: the ordered document body (blueprint sections and assets, per the layout). */
    public List<ChapterView> chapters = new ArrayList<>();
    /** Plan export only: every exported asset, in document order (e.g. for an overview table). */
    public List<AssetView> assets = new ArrayList<>();
    public LayoutView layout;
    /** Interface texts in the plan's language: {@code {{labels.version}}}. */
    public Map<String, String> labels = new LinkedHashMap<>();
    public ExportInfo export;
    public BrandingView branding;
    public TableOfContents toc = new TableOfContents();
    public PageBreak pageBreak = new PageBreak();

    public static class PlanView {
        public String id;
        public String title;
        /** Whether the title is long enough for the cover to set it smaller; {@code shortTitle} is its opposite. */
        public boolean longTitle;
        public boolean shortTitle;
        /** What kind of plan this is (the blueprint type, e.g. "Data Management Plan"). */
        public String kind;
        public String version;
        public String status;
        public boolean finalized;
        public RichText description;
        public String blueprint;
        public String language;
        public String access;
        public String doi;
        public String createdAt;
        public String updatedAt;
        public String finalizedAt;
        public String creator;
        public List<PersonView> contacts = new ArrayList<>();
        public List<PersonView> users = new ArrayList<>();
        /** Plan references by reference type code: loop with {@code {{?plan.refs.grants}}{{label}}{{/plan.refs.grants}}} or render with {@code {{>plan.refs.grants}}}. */
        public Map<String, List<ReferenceView>> refs = new LinkedHashMap<>();
        /** Answers by field semantic, dots replaced by underscores: {@code {{plan.sem.rda_dmp_ethical_issues_description}}}. */
        public Map<String, String> sem = new LinkedHashMap<>();
        /** Every exported blueprint question, flat and in order. */
        public List<QuestionView> questions = new ArrayList<>();
        /** Key facts for a cover page (references, contacts, DOI) as label/value pairs. */
        public Facts facts;
    }

    /** One entry of the plan body: exactly one of {@code section}, {@code asset}, {@code group} is set. */
    public static class ChapterView {
        public SectionView section;
        public AssetView asset;
        public GroupView group;
    }

    /** A heading introducing a group of assets (asset placement "before" or "after"). */
    public static class GroupView {
        public String title;
        /** Set when headings are numbered: the Assets chapter of grouped placements is numbered like a section. */
        public String number;
        public String numberPrefix;
        public String heading;
        /** True when the layout starts sections on a new page (never for the first chapter). */
        public boolean newPage;
    }

    /** A root section of the plan blueprint. */
    public static class SectionView {
        public String id;
        public String number;
        public String title;
        /** Number and title, ready to print; just the title when numbering is off. */
        public String heading;
        /** The number followed by a space, or absent when numbering is off: {@code {{numberPrefix}}{{title}}}. */
        public String numberPrefix;
        /** The section's guidance text; only set when the layout exports guidance. */
        public RichText description;
        /** True when the layout starts sections on a new page (never for the first chapter). */
        public boolean newPage;
        public Body body;
        public boolean hasBody;
        /** The assets of this section, in document order (whatever the placement). */
        public List<AssetView> assets = new ArrayList<>();
        public List<QuestionView> questions = new ArrayList<>();
    }

    public static class AssetView {
        public String id;
        public String number;
        public String title;
        /** The number and an en space, for templates that style the number apart (like section headings). */
        public String numberPrefix;
        /** Number and title, e.g. "8.1 Survey dataset"; the title alone when headings are not numbered. */
        public String heading;
        /** Whether the title is long enough for the cover of an asset export to set it smaller. */
        public boolean longTitle;
        public boolean shortTitle;
        public RichText description;
        public String template;
        public String type;
        public String status;
        public boolean finalized;
        public String section;
        public String createdAt;
        public String tags;
        /** Template, type and status joined, for a one-line subtitle. */
        public String meta;
        /** True when the layout asks for each asset to start on a new page. */
        public boolean newPage;
        public Body body;
        public Map<String, List<ReferenceView>> refs = new LinkedHashMap<>();
        public Map<String, String> sem = new LinkedHashMap<>();
        public List<QuestionView> questions = new ArrayList<>();
        public Facts facts;
    }

    public static class QuestionView {
        public String label;
        public String section;
        /** The answer as plain text (lists joined with ", "). */
        public String answer;
        public List<String> semantics = new ArrayList<>();
    }

    public static class ReferenceView {
        public String label;
        public String reference;
        public String url;
        public String type;
        public String typeCode;
        /** The persistent identifier found in the label or fields, e.g. an ORCID. */
        public String pid;
        public String pidType;
        /** Every field value by code (non-alphanumerics replaced by underscores). */
        public Map<String, String> fields = new LinkedHashMap<>();
        /** The fields selected for display, labelled, in order. */
        public List<FactView> details = new ArrayList<>();
        /** The details joined into one line. */
        public String detailsText;
    }

    public static class FactView {
        public String label;
        public String value;
        public String url;

        public FactView() {
        }

        public FactView(String label, String value, String url) {
            this.label = label;
            this.value = value;
            this.url = url;
        }
    }

    public static class PersonView {
        public String name;
        public String email;
        public String role;
    }

    public static class LayoutView {
        public boolean cover;
        public boolean toc;
        public boolean numbering;
        public boolean assetsInSections;
        public boolean assetsBeforeSections;
        public boolean assetsAfterSections;
        public boolean assetPageBreak;
        public String questionStyle;
        public String referenceStyle;
        public String tableStyle;
    }

    public static class ExportInfo {
        public String date;
        public String generator;
        /** The generator and export date as one line, e.g. "Generated with OpenCDMP · Exported 25/09/2026". */
        public String note;
    }

    public static class BrandingView {
        /** The configured cover logo, rendered with {@code {{@branding.logo}}}; absent when none is configured. */
        public PictureRenderData logo;
        public String accentColor;
    }
}
