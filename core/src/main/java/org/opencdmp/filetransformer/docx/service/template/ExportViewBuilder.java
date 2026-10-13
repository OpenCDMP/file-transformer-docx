package org.opencdmp.filetransformer.docx.service.template;

import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import org.opencdmp.commonmodels.enums.DescriptionStatus;
import org.opencdmp.commonmodels.enums.FieldType;
import org.opencdmp.commonmodels.enums.PlanAccessType;
import org.opencdmp.commonmodels.enums.PlanStatus;
import org.opencdmp.commonmodels.enums.SystemFieldType;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.PlanUserModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.plan.PlanContactModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.commonmodels.models.planblueprint.RootSectionModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.FieldModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.FieldSetModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.SectionModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.ReferenceTypeDataModel;
import org.opencdmp.commonmodels.models.planblueprintdescriptiontemplatefields.fielddata.SystemFieldDataModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.PropertyDefinitionFieldSetItemModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.PropertyDefinitionFieldSetModel;
import org.opencdmp.commonmodels.models.plandescriptionfields.VisibilityStateModel;
import org.opencdmp.commonmodels.models.planreference.PlanReferenceModel;
import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.service.language.LanguageService;
import org.opencdmp.filetransformer.docx.service.pid.PidService;
import org.opencdmp.filetransformer.docx.service.storage.FileStorageService;
import org.opencdmp.filetransformer.docx.service.template.content.Body;
import org.opencdmp.filetransformer.docx.service.template.content.Facts;
import org.opencdmp.filetransformer.docx.service.template.content.RichText;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.AssetView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ChapterView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.FactView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.GroupView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.PersonView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.PlanView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.QuestionView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.ReferenceView;
import org.opencdmp.filetransformer.docx.service.template.view.ExportView.SectionView;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordExportOptions;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.io.ByteArrayInputStream;
import java.time.Instant;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Builds the {@link ExportView} a '{{...}}' template renders against. This is where every
 * decision about <i>what</i> is exported happens — visibility, "include in export", selective
 * export, unanswered questions, asset placement, numbering — so templates only decide
 * <i>how</i> it looks.
 */
@Component
public class ExportViewBuilder {
    private static final Logger logger = LoggerFactory.getLogger(ExportViewBuilder.class);

    /** Heading level of a plan section's sub-sections (the section itself is Heading 1). */
    private static final int PLAN_SECTION_BODY_LEVEL = 2;
    /** Heading level of an asset's sections inside a plan export (the asset title sits at level 2). */
    private static final int PLAN_ASSET_BODY_LEVEL = 3;
    /** Heading level of an asset's sections in an asset export. */
    private static final int ASSET_BODY_LEVEL = 1;
    private static final double LOGO_MAX_WIDTH_PX = 240;
    private static final double LOGO_MAX_HEIGHT_PX = 64;

    private final PidService pidService;
    private final LanguageService languageService;
    private final FileStorageService fileStorageService;
    private final WordFileTransformerServiceProperties properties;

    public ExportViewBuilder(PidService pidService, LanguageService languageService, FileStorageService fileStorageService, WordFileTransformerServiceProperties properties) {
        this.pidService = pidService;
        this.languageService = languageService;
        this.fileStorageService = fileStorageService;
        this.properties = properties;
    }

    public ExportView forPlan(PlanModel plan, WordExportOptions options) {
        Run run = new Run(plan, options);
        ExportView view = run.base();
        view.plan = run.planView();

        Visibility visibility = new Visibility(plan.getVisibilityStates());
        Map<String, PropertyDefinitionFieldSetModel> values = plan.getProperties() != null ? plan.getProperties().getFieldSets() : null;
        // cover facts describe the whole plan: sections left out of the export still contribute
        Collector everything = new Collector(null);

        // first the sections that export anything, so numbers can follow where the assets end up
        record Exported(RootSectionModel section, boolean visible, boolean hasBody, List<DescriptionModel> assets) {
        }
        List<Exported> exported = new ArrayList<>();
        for (RootSectionModel section : sortedSections(plan)) {
            // the platform reports a section without visible questions as hidden, which includes a section that only
            // holds assets; its assets are still part of the plan, so only its questions are left out
            boolean sectionVisible = visibility.isVisible(section.getId(), null, true);
            List<DescriptionModel> assets = run.assetsOf(section);
            if (!sectionVisible && assets.isEmpty()) continue;
            if (!options.isSectionIncluded(section.getId())) {
                run.nodes(section.getSections(), section.getFieldSets(), values, visibility, null, null, everything.within(section.getTitle()));
                continue;
            }
            boolean hasBody = sectionVisible && !run.nodes(section.getSections(), section.getFieldSets(), values, visibility, null, null, new Collector(section.getTitle())).isEmpty();
            if (!hasBody && assets.isEmpty()) continue;
            exported.add(new Exported(section, sectionVisible, hasBody, assets));
        }

        // Numbering: "after their section" makes each asset a sub-chapter of its section (8.1, then 8.1.1 inside it);
        // grouped placements number the Assets chapter like a section, before or after the others.
        WordExportOptions.AssetPlacement placement = options.getAssetPlacement();
        boolean grouped = placement != WordExportOptions.AssetPlacement.SECTIONS;
        boolean anyAssets = exported.stream().anyMatch(x -> !x.assets().isEmpty());
        int sectionNumber = grouped && placement == WordExportOptions.AssetPlacement.BEFORE && anyAssets ? 1 : 0;
        List<SectionView> sections = new ArrayList<>();
        List<DescriptionModel> groupedAssets = new ArrayList<>();
        List<String> groupedAssetSections = new ArrayList<>();
        for (Exported item : exported) {
            RootSectionModel section = item.section();
            // in grouped placements a section left with only assets is not a chapter of its own
            if (grouped && !item.hasBody()) {
                item.assets().forEach(x -> { groupedAssets.add(x); groupedAssetSections.add(section.getTitle()); });
                continue;
            }
            sectionNumber++;
            String number = String.valueOf(sectionNumber);
            Collector collector = new Collector(section.getTitle());
            List<Body.Node> nodes = item.visible() ? run.nodes(section.getSections(), section.getFieldSets(), values,
                    visibility, null, options.isNumbering() ? number : null, collector) : new ArrayList<>();
            everything.absorb(collector);

            List<AssetView> assets = new ArrayList<>();
            if (grouped) {
                item.assets().forEach(x -> { groupedAssets.add(x); groupedAssetSections.add(section.getTitle()); });
            } else {
                for (int i = 0; i < item.assets().size(); i++) {
                    // after the section's own sub-sections and question groups: 3.1-3.4 are questions, 3.5 the first asset
                    assets.add(run.assetView(item.assets().get(i), childNumber(number, nodes.size() + i + 1), PLAN_ASSET_BODY_LEVEL, section.getTitle()));
                }
            }

            SectionView sectionView = new SectionView();
            sectionView.id = section.getId();
            sectionView.number = options.isNumbering() ? number : null;
            sectionView.title = section.getTitle();
            sectionView.heading = options.isNumbering() ? number + "\u2002" + nonNull(section.getTitle()) : section.getTitle();
            sectionView.numberPrefix = options.isNumbering() ? number + "\u2002" : null;
            sectionView.description = options.isGuidance() ? richText(section.getDescription()) : null;
            sectionView.newPage = options.isSectionPageBreak();
            sectionView.body = new Body(PLAN_SECTION_BODY_LEVEL, nodes);
            sectionView.hasBody = !nodes.isEmpty();
            sectionView.assets = assets;
            sectionView.questions = collector.questions;
            view.plan.questions.addAll(collector.questions);
            collector.semantics.forEach((k, v) -> view.plan.sem.merge(k, v, (a, b) -> a + "; " + b));
            sections.add(sectionView);
        }

        // the Assets chapter of grouped placements: first when before, next number when after
        String groupNumber = String.valueOf(placement == WordExportOptions.AssetPlacement.BEFORE ? 1 : sectionNumber + 1);
        List<AssetView> allAssets = new ArrayList<>();
        if (grouped) {
            for (int i = 0; i < groupedAssets.size(); i++) {
                allAssets.add(run.assetView(groupedAssets.get(i), childNumber(groupNumber, i + 1), PLAN_ASSET_BODY_LEVEL, groupedAssetSections.get(i)));
            }
        } else {
            sections.forEach(x -> allAssets.addAll(x.assets));
        }

        view.plan.facts = new Facts(run.planFacts(everything));

        switch (placement) {
            case SECTIONS -> sections.forEach(section -> {
                view.chapters.add(chapter(section));
                section.assets.forEach(asset -> view.chapters.add(chapter(asset)));
            });
            case BEFORE -> {
                if (!allAssets.isEmpty()) view.chapters.add(group(run.labels.get("assets"), options.isNumbering() ? groupNumber : null));
                allAssets.forEach(asset -> view.chapters.add(chapter(asset)));
                sections.forEach(section -> view.chapters.add(chapter(section)));
            }
            case AFTER -> {
                sections.forEach(section -> view.chapters.add(chapter(section)));
                if (!allAssets.isEmpty()) view.chapters.add(group(run.labels.get("assets"), options.isNumbering() ? groupNumber : null));
                allAssets.forEach(asset -> view.chapters.add(chapter(asset)));
            }
        }
        for (int i = 0; i < view.chapters.size(); i++) {
            ChapterView chapter = view.chapters.get(i);
            boolean first = i == 0 || view.chapters.get(i - 1).group != null;
            if (chapter.group != null) chapter.group.newPage = i > 0 && options.isSectionPageBreak();
            if (chapter.section != null && first) chapter.section.newPage = false;
            if (chapter.asset != null && first) chapter.asset.newPage = false;
        }
        view.assets = allAssets;
        return view;
    }

    public ExportView forAsset(DescriptionModel asset, WordExportOptions options) {
        Run run = new Run(asset.getPlan(), options);
        ExportView view = run.base();
        view.plan = run.planView();
        view.plan.facts = new Facts(run.planFacts(null));
        String sectionTitle = null;
        if (asset.getPlan() != null && asset.getSectionId() != null) {
            sectionTitle = sortedSections(asset.getPlan()).stream().filter(x -> asset.getSectionId().equals(x.getId())).map(SectionModel::getTitle).findFirst().orElse(null);
        }
        view.asset = run.assetView(asset, null, ASSET_BODY_LEVEL, sectionTitle);
        view.asset.newPage = false;
        view.assets = List.of(view.asset);
        return view;
    }

    // ------------------------------------------------------------------ per export state

    private final class Run {
        private final PlanModel plan;
        private final WordExportOptions options;
        private final ExportLabels labels;
        private final DateTimeFormatter dateFormatter;
        private final AnswerResolver answers;
        private final ReferenceFormatter references;

        Run(PlanModel plan, WordExportOptions options) {
            this.plan = plan;
            this.options = options;
            this.labels = ExportLabels.forLanguage(plan != null ? plan.getLanguage() : null).withOverrides(options.getLabelOverrides());
            Locale locale = plan != null && !isBlank(plan.getLanguage()) ? Locale.forLanguageTag(plan.getLanguage().trim()) : Locale.ENGLISH;
            this.dateFormatter = DateTimeFormatter.ofPattern(options.getDateFormat(), locale).withZone(ZoneOffset.UTC); // date answers are stored as UTC midnight (MomentUtcDateAdapter)
            Map<UUID, String> assetTitles = new HashMap<>();
            if (plan != null && plan.getDescriptions() != null) {
                plan.getDescriptions().stream().filter(x -> x.getId() != null && x.getLabel() != null).forEach(x -> assetTitles.put(x.getId(), x.getLabel()));
            }
            this.references = new ReferenceFormatter(pidService, options.getReferenceFields(), this.dateFormatter);
            this.answers = new AnswerResolver(this.labels, this.dateFormatter, this.references, languageService, fileStorageService, properties.isUseSharedStorage(), plan, assetTitles);
        }

        ExportView base() {
            ExportView view = new ExportView();
            view.labels = this.labels.asMap();
            ExportView.LayoutView layout = new ExportView.LayoutView();
            layout.cover = this.options.isCover();
            layout.toc = this.options.isToc();
            layout.numbering = this.options.isNumbering();
            layout.assetsInSections = this.options.getAssetPlacement() == WordExportOptions.AssetPlacement.SECTIONS;
            layout.assetsBeforeSections = this.options.getAssetPlacement() == WordExportOptions.AssetPlacement.BEFORE;
            layout.assetsAfterSections = this.options.getAssetPlacement() == WordExportOptions.AssetPlacement.AFTER;
            layout.assetPageBreak = this.options.isAssetPageBreak();
            layout.questionStyle = this.options.getQuestionStyle().name().toLowerCase(Locale.ROOT);
            layout.referenceStyle = this.options.getReferenceStyle().name().toLowerCase(Locale.ROOT);
            layout.tableStyle = this.options.getTableStyle().name().toLowerCase(Locale.ROOT);
            view.layout = layout;
            view.export = new ExportView.ExportInfo();
            view.export.date = this.dateFormatter.format(Instant.now());
            view.export.generator = isBlank(this.labels.get("generatedWith")) ? null : this.labels.get("generatedWith");
            view.export.note = Views.join("  \u00B7  ", view.export.generator, Views.join(" ", this.labels.get("exported"), view.export.date));
            view.branding = new ExportView.BrandingView();
            view.branding.logo = this.logo(this.options.getCoverLogo());
            view.branding.accentColor = this.options.getAccentColor();
            return view;
        }

        private PlanView planView;

        private PlanView planViewCache() {
            if (this.planView == null) this.planView = this.planView();
            return this.planView;
        }

        PlanView planView() {
            if (this.planView != null) return this.planView;
            PlanView view = new PlanView();
            this.planView = view;
            if (this.plan == null) return view;
            view.id = this.plan.getId() != null ? this.plan.getId().toString() : null;
            view.title = this.plan.getLabel();
            view.longTitle = Views.isLongTitle(view.title);
            view.shortTitle = !view.longTitle;
            view.kind = !this.options.getLabelOverrides().containsKey("planKind") && this.plan.getPlanBlueprint() != null && this.plan.getPlanBlueprint().getType() != null
                    && !isBlank(this.plan.getPlanBlueprint().getType().getName()) ? this.plan.getPlanBlueprint().getType().getName() : this.labels.get("planKind");
            view.version = String.valueOf(this.plan.getVersion());
            view.status = this.plan.getStatus() != null ? this.plan.getStatus().getName() : null;
            view.finalized = this.plan.getStatus() != null && this.plan.getStatus().getInternalStatus() == PlanStatus.Finalized;
            view.description = richText(this.plan.getDescription());
            view.blueprint = this.plan.getPlanBlueprint() != null ? this.plan.getPlanBlueprint().getLabel() : null;
            view.language = this.languageName(this.plan.getLanguage());
            view.access = this.plan.getAccessType() != null ? this.labels.get(this.plan.getAccessType() == PlanAccessType.Public ? "public" : "restricted") : null;
            view.doi = this.plan.getEntityDois() != null && !this.plan.getEntityDois().isEmpty() ? this.plan.getEntityDois().getFirst().getDoi() : null;
            view.createdAt = this.date(this.plan.getCreatedAt());
            view.updatedAt = this.date(this.plan.getUpdatedAt());
            view.finalizedAt = this.date(this.plan.getFinalizedAt());
            view.creator = this.plan.getCreator() != null ? this.plan.getCreator().getName() : null;
            if (this.plan.getProperties() != null && this.plan.getProperties().getContacts() != null) {
                for (PlanContactModel contact : this.plan.getProperties().getContacts()) {
                    if (contact == null) continue;
                    PersonView person = new PersonView();
                    person.name = Views.join(" ", contact.getFirstName(), contact.getLastName());
                    person.email = contact.getEmail();
                    view.contacts.add(person);
                }
            }
            if (this.plan.getUsers() != null) {
                for (PlanUserModel user : this.plan.getUsers()) {
                    if (user == null || user.getUser() == null) continue;
                    PersonView person = new PersonView();
                    person.name = user.getUser().getName();
                    person.role = user.getRole() != null ? Views.humanize(user.getRole().name()) : null;
                    view.users.add(person);
                }
            }
            if (this.plan.getReferences() != null) {
                for (PlanReferenceModel planReference : this.plan.getReferences()) {
                    if (planReference.getReference() == null || planReference.getReference().getType() == null || planReference.getReference().getType().getCode() == null) continue;
                    view.refs.computeIfAbsent(Views.key(planReference.getReference().getType().getCode()), k -> new ArrayList<>()).add(this.references.format(planReference.getReference()));
                }
            }
            return view;
        }

        /**
         * The plan cover facts, per the {@code layout-plan-facts} setting (default: funder, grant,
         * researchers, organisations, contact, DOI). Rows without a value are left out.
         *
         * @param answers the answers of the whole plan, for {@code sem:} rows; {@code null} in asset exports
         */
        List<FactView> planFacts(Collector answers) {
            List<FactView> facts = new ArrayList<>();
            if (this.plan == null) return facts;
            List<WordExportOptions.FactSpec> specs = this.options.getPlanFacts() != null ? this.options.getPlanFacts() : List.of(
                    new WordExportOptions.FactSpec("funders", null),
                    new WordExportOptions.FactSpec(properties.getGrantReferenceCode(), null),
                    new WordExportOptions.FactSpec(properties.getResearcherReferenceCode(), null),
                    new WordExportOptions.FactSpec(properties.getOrganizationReferenceCode(), null),
                    new WordExportOptions.FactSpec("contacts", null),
                    new WordExportOptions.FactSpec("doi", null));
            for (WordExportOptions.FactSpec spec : specs) {
                if (spec.key() == null) continue;
                FactView fact = this.planFact(spec.key().trim(), answers);
                if (fact == null || isBlank(fact.value)) continue;
                if (spec.label() != null) fact.label = spec.label();
                facts.add(fact);
            }
            return facts;
        }

        private FactView planFact(String key, Collector answers) {
            PlanView plan = this.planViewCache();
            switch (key.toLowerCase(Locale.ROOT)) {
                case "contacts" -> {
                    List<String> contacts = AnswerResolver.contacts(this.plan.getProperties() != null ? this.plan.getProperties().getContacts() : null);
                    return new FactView(this.labels.get("contacts"), String.join(", ", contacts), null);
                }
                case "doi" -> {
                    return plan.doi != null ? new FactView(this.labels.get("doi"), plan.doi, "https://doi.org/" + plan.doi) : null;
                }
                case "version" -> {
                    return new FactView(this.labels.get("version"), plan.version, null);
                }
                case "status" -> {
                    return new FactView(this.labels.get("status"), plan.status, null);
                }
                case "blueprint" -> {
                    return new FactView(this.labels.get("blueprint"), plan.blueprint, null);
                }
                case "language" -> {
                    return new FactView(this.labels.get("language"), plan.language, null);
                }
                case "access" -> {
                    return new FactView(this.labels.get("access"), plan.access, null);
                }
                case "created" -> {
                    return new FactView(this.labels.get("created"), plan.createdAt, null);
                }
                case "updated" -> {
                    return new FactView(this.labels.get("updated"), plan.updatedAt, null);
                }
                case "finalized" -> {
                    return new FactView(this.labels.get("finalized"), plan.finalizedAt, null);
                }
                case "creator" -> {
                    return new FactView(this.labels.get("creator"), plan.creator, null);
                }
                case "members" -> {
                    return new FactView(this.labels.get("users"), String.join(", ", AnswerResolver.users(this.plan.getUsers())), null);
                }
                default -> {
                }
            }
            if (key.startsWith("sem:")) return answers != null ? answers.semanticFact(key.substring(4)) : null;

            // a reference type code: the plan's references of that type, labelled like the blueprint question asking for them
            List<String> labels = new ArrayList<>();
            String typeName = null;
            if (this.plan.getReferences() != null) {
                Set<UUID> seen = new HashSet<>();
                for (PlanReferenceModel planReference : this.plan.getReferences()) {
                    var reference = planReference.getReference();
                    if (reference == null || reference.getType() == null || !key.equalsIgnoreCase(reference.getType().getCode())) continue;
                    if (reference.getId() != null && !seen.add(reference.getId())) continue;
                    typeName = reference.getType().getName();
                    labels.add(this.references.format(reference).label);
                }
            }
            if (labels.isEmpty()) return null;
            String label = this.blueprintQuestionLabel(key);
            return new FactView(label != null ? label : typeName != null ? typeName : Views.humanize(key), String.join(", ", labels), null);
        }

        /** The label of the first blueprint question asking for references of a type. */
        private String blueprintQuestionLabel(String typeCode) {
            if (this.plan.getPlanBlueprint() == null || this.plan.getPlanBlueprint().getDefinition() == null) return null;
            for (FieldModel field : this.plan.getPlanBlueprint().getDefinition().getAllField()) {
                if (field.getData() instanceof ReferenceTypeDataModel reference && reference.getReferenceType() != null
                        && typeCode.equalsIgnoreCase(reference.getReferenceType().getCode())) {
                    String label = AnswerResolver.label(field);
                    if (label != null) return label;
                }
            }
            return null;
        }

        /**
         * The asset cover facts, per the {@code layout-asset-facts} setting (default: type, template,
         * plan, section, status, tags). Besides those keys, a row can be a reference type code (the
         * asset's references of that type) or {@code sem:<semantic>}.
         */
        List<FactView> assetFacts(AssetView asset, Collector answers) {
            List<WordExportOptions.FactSpec> specs = this.options.getAssetFacts() != null ? this.options.getAssetFacts() : List.of(
                    new WordExportOptions.FactSpec("type", null), new WordExportOptions.FactSpec("template", null),
                    new WordExportOptions.FactSpec("plan", null), new WordExportOptions.FactSpec("section", null),
                    new WordExportOptions.FactSpec("status", null), new WordExportOptions.FactSpec("tags", null));
            List<FactView> facts = new ArrayList<>();
            for (WordExportOptions.FactSpec spec : specs) {
                if (spec.key() == null) continue;
                String key = spec.key().trim();
                FactView fact = switch (key.toLowerCase(Locale.ROOT)) {
                    case "type" -> new FactView(this.labels.get("type"), asset.type, null);
                    case "template" -> new FactView(this.labels.get("template"), asset.template, null);
                    case "plan" -> new FactView(this.labels.get("partOf"), this.plan != null ? this.plan.getLabel() : null, null);
                    case "section" -> new FactView(this.labels.get("section"), asset.section, null);
                    case "status" -> new FactView(this.labels.get("status"), asset.status, null);
                    case "tags" -> new FactView(this.labels.get("tags"), asset.tags, null);
                    case "created" -> new FactView(this.labels.get("created"), asset.createdAt, null);
                    default -> {
                        if (key.startsWith("sem:")) yield answers.semanticFact(key.substring(4));
                        List<ReferenceView> references = asset.refs.get(Views.key(key));
                        if (references == null || references.isEmpty()) yield null;
                        yield new FactView(references.getFirst().type != null ? references.getFirst().type : Views.humanize(key),
                                String.join(", ", references.stream().map(x -> x.label).toList()), null);
                    }
                };
                if (fact == null || isBlank(fact.value)) continue;
                if (spec.label() != null) fact.label = spec.label();
                facts.add(fact);
            }
            return facts;
        }

        /** The plan's assets exported under a blueprint section: not cancelled, not excluded, finalized only for finalized / public plans. */
        List<DescriptionModel> assetsOf(RootSectionModel section) {
            if (this.plan.getDescriptions() == null) return List.of();
            boolean finalizedPlan = this.plan.getStatus() != null && this.plan.getStatus().getInternalStatus() == PlanStatus.Finalized;
            boolean publicPlan = this.plan.getPublicAfter() != null && this.plan.getPublicAfter().isAfter(Instant.now());
            return this.plan.getDescriptions().stream()
                    .filter(x -> x.getStatus() != null && x.getStatus().getInternalStatus() != DescriptionStatus.Canceled)
                    .filter(x -> !(finalizedPlan || publicPlan) || x.getStatus().getInternalStatus() == DescriptionStatus.Finalized)
                    .filter(x -> section.getId() != null && section.getId().equals(x.getSectionId()))
                    .filter(x -> !this.options.isDescriptionExcluded(x.getId()))
                    .sorted(Comparator.comparing(DescriptionModel::getCreatedAt, Comparator.nullsLast(Comparator.naturalOrder())))
                    .toList();
        }

        AssetView assetView(DescriptionModel asset, String number, int bodyLevel, String sectionTitle) {
            AssetView view = new AssetView();
            view.id = asset.getId() != null ? asset.getId().toString() : null;
            boolean numbered = this.options.isNumbering() && number != null;
            view.number = numbered ? number : null;
            view.title = asset.getLabel();
            view.numberPrefix = numbered ? number + "\u2002" : null;
            view.heading = numbered ? number + "\u2002" + nonNull(asset.getLabel()) : asset.getLabel();
            view.description = richText(asset.getDescription());
            view.template = asset.getDescriptionTemplate() != null ? asset.getDescriptionTemplate().getLabel() : null;
            view.type = asset.getDescriptionTemplate() != null && asset.getDescriptionTemplate().getType() != null ? asset.getDescriptionTemplate().getType().getName() : null;
            view.status = asset.getStatus() != null ? asset.getStatus().getName() : null;
            view.finalized = asset.getStatus() != null && asset.getStatus().getInternalStatus() == DescriptionStatus.Finalized;
            view.section = sectionTitle;
            view.createdAt = this.date(asset.getCreatedAt());
            view.tags = asset.getTags() != null && !asset.getTags().isEmpty() ? String.join(", ", asset.getTags()) : null;
            // "Dataset · Dataset (Horizon Europe) · Draft" reads as a stutter: the template name already says the type
            boolean typeInTemplate = view.type != null && view.template != null && view.template.toLowerCase(Locale.ROOT).contains(view.type.toLowerCase(Locale.ROOT));
            view.meta = Views.join("  ·  ", typeInTemplate ? null : view.type, view.template, view.status);
            view.longTitle = Views.isLongTitle(view.title);
            view.shortTitle = !view.longTitle;
            view.newPage = this.options.isAssetPageBreak();

            Collector collector = new Collector(null);
            List<Body.Node> nodes = List.of();
            if (asset.getDescriptionTemplate() != null && asset.getDescriptionTemplate().getDefinition() != null) {
                // inside a plan an asset's sections continue its number (8.1 -> 8.1.1); on their own they start at 1
                nodes = this.nodes(asset.getDescriptionTemplate().getDefinition().getSections(), null,
                        asset.getProperties() != null ? asset.getProperties().getFieldSets() : null,
                        new Visibility(asset.getVisibilityStates()), asset, numbered ? number : null, collector);
            }
            view.body = new Body(bodyLevel, nodes);
            view.questions = collector.questions;
            view.sem = collector.semantics;
            view.refs = collector.references;

            view.facts = new Facts(this.assetFacts(view, collector));
            return view;
        }

        /**
         * Sub-sections then field sets of one level, numbered in the order they are exported;
         * a level with nothing to export yields no node at all.
         */
        List<Body.Node> nodes(List<? extends SectionModel> sections, List<FieldSetModel> fieldSets, Map<String, PropertyDefinitionFieldSetModel> values,
                              Visibility visibility, DescriptionModel asset, String prefix, Collector collector) {
            List<Body.Node> nodes = new ArrayList<>();
            int counter = 0;
            if (sections != null) {
                for (SectionModel section : sections.stream().sorted(Comparator.comparing(SectionModel::getOrdinal, Comparator.nullsLast(Comparator.naturalOrder()))).toList()) {
                    if (!visibility.isVisible(section.getId(), null, true)) continue;
                    String number = this.options.isNumbering() ? childNumber(prefix, counter + 1) : null;
                    List<Body.Node> children = this.nodes(section.getSections(), section.getFieldSets(), values, visibility, asset, number, collector.within(section.getTitle()));
                    if (children.isEmpty()) continue;
                    counter++;
                    nodes.add(new Body.SectionNode(number, section.getTitle(), this.options.isGuidance() ? richText(section.getDescription()) : null, children));
                }
            }
            if (fieldSets != null) {
                for (FieldSetModel fieldSet : fieldSets.stream().sorted(Comparator.comparing(FieldSetModel::getOrdinal, Comparator.nullsLast(Comparator.naturalOrder()))).toList()) {
                    String number = this.options.isNumbering() ? childNumber(prefix, counter + 1) : null;
                    Body.FieldSetNode node = this.fieldSet(fieldSet, values != null ? values.get(fieldSet.getId()) : null, visibility, asset, number, collector);
                    if (node == null) continue;
                    counter++;
                    nodes.add(node);
                }
            }
            return nodes;
        }

        private Body.FieldSetNode fieldSet(FieldSetModel fieldSet, PropertyDefinitionFieldSetModel stored, Visibility visibility, DescriptionModel asset, String number, Collector collector) {
            List<FieldModel> fields = fieldSet.getFields() == null ? List.of() : fieldSet.getFields().stream()
                    .filter(x -> !Boolean.FALSE.equals(x.getIncludeInExport()))
                    .filter(x -> !this.isShownElsewhere(x))
                    .sorted(Comparator.comparing(FieldModel::getOrdinal, Comparator.nullsLast(Comparator.naturalOrder())))
                    .toList();
            if (fields.isEmpty()) return null;

            boolean placeholders = this.options.getUnansweredQuestions() == WordExportOptions.UnansweredQuestionsMode.PLACEHOLDER;
            boolean hasMultiplicity = fieldSet.getMultiplicity() != null;
            Body.Mode mode;
            if (hasMultiplicity && Boolean.TRUE.equals(fieldSet.getMultiplicity().getTransposedTableView())) mode = Body.Mode.TRANSPOSED;
            else if (hasMultiplicity && Boolean.TRUE.equals(fieldSet.getMultiplicity().getTableView())) mode = Body.Mode.TABLE;
            else mode = switch (this.options.getQuestionStyle()) {
                case STACKED -> Body.Mode.STACKED;
                case TABLE -> Body.Mode.QUESTION_TABLE;
                case COMPACT -> Body.Mode.COMPACT;
            };
            boolean tabular = mode == Body.Mode.TABLE || mode == Body.Mode.TRANSPOSED;

            // system fields store no values: they get a synthetic first item
            List<PropertyDefinitionFieldSetItemModel> items = stored != null && stored.getItems() != null
                    ? stored.getItems().stream().sorted(Comparator.comparing(PropertyDefinitionFieldSetItemModel::getOrdinal, Comparator.nullsLast(Comparator.naturalOrder()))).toList()
                    : List.of();
            boolean synthetic = false;
            if (items.isEmpty() && (placeholders || fields.stream().anyMatch(x -> x.getData() != null && x.getData().getFieldType() == FieldType.SYSTEM))) {
                PropertyDefinitionFieldSetItemModel item = new PropertyDefinitionFieldSetItemModel();
                item.setOrdinal(0);
                item.setFields(Map.of());
                items = List.of(item);
                synthetic = true;
            }

            List<List<FieldModel>> itemFields = new ArrayList<>();
            List<List<Body.Answer>> itemAnswers = new ArrayList<>();
            for (PropertyDefinitionFieldSetItemModel item : items) {
                if (!visibility.isVisible(fieldSet.getId(), item.getOrdinal(), synthetic)) continue;
                List<FieldModel> shown = new ArrayList<>();
                List<Body.Answer> answers = new ArrayList<>();
                boolean answered = false;
                for (FieldModel field : fields) {
                    if (!visibility.isVisible(field.getId(), item.getOrdinal(), synthetic)) continue;
                    Body.Answer answer;
                    try {
                        answer = this.answers.resolve(field, item.getFields() != null ? item.getFields().get(field.getId()) : null, asset);
                    } catch (Exception e) {
                        logger.warn("could not resolve the answer of field {}, exporting it as unanswered", field.getId(), e);
                        answer = null;
                    }
                    if (answer == null && !placeholders && !tabular) continue;
                    answered = answered || answer != null;
                    shown.add(field);
                    answers.add(answer != null ? answer : new Body.Missing());
                }
                if (!answered && !placeholders) continue;
                itemFields.add(shown);
                itemAnswers.add(answers);
            }
            if (itemAnswers.isEmpty()) return null;

            String title = fieldSet.getTitle();
            List<String> columns = new ArrayList<>();
            List<Body.Item> result = new ArrayList<>();
            if (tabular) {
                // one column per question answered in at least one item, in field order
                List<FieldModel> columnFields = fields.stream().filter(field -> {
                    for (int i = 0; i < itemFields.size(); i++) {
                        int index = itemFields.get(i).indexOf(field);
                        if (index >= 0 && (placeholders || !(itemAnswers.get(i).get(index) instanceof Body.Missing))) return true;
                    }
                    return false;
                }).toList();
                columnFields.forEach(field -> columns.add(nonNull(AnswerResolver.label(field))));
                for (int i = 0; i < itemFields.size(); i++) {
                    List<Body.Question> questions = new ArrayList<>();
                    for (FieldModel field : columnFields) {
                        int index = itemFields.get(i).indexOf(field);
                        questions.add(new Body.Question(AnswerResolver.label(field), index >= 0 ? itemAnswers.get(i).get(index) : new Body.Missing()));
                    }
                    result.add(new Body.Item(questions));
                }
            } else {
                for (int i = 0; i < itemFields.size(); i++) {
                    List<Body.Question> questions = new ArrayList<>();
                    for (int j = 0; j < itemFields.get(i).size(); j++) {
                        String label = AnswerResolver.label(itemFields.get(i).get(j));
                        // a lone question repeating its field set title needs no label of its own
                        if (label != null && title != null && label.equalsIgnoreCase(title.trim())) label = null;
                        questions.add(new Body.Question(label, itemAnswers.get(i).get(j)));
                    }
                    result.add(new Body.Item(questions));
                }
            }

            for (int i = 0; i < itemFields.size(); i++) {
                for (int j = 0; j < itemFields.get(i).size(); j++) {
                    collector.add(itemFields.get(i).get(j), title, itemAnswers.get(i).get(j));
                }
            }
            String comment = this.options.isComments() && stored != null && stored.getComment() != null && !stored.getComment().isBlank() ? stored.getComment() : null;
            RichText guidance = this.options.isGuidance() ? richText(Views.join("", fieldSet.getDescription(), fieldSet.getExtendedDescription())) : null;
            return new Body.FieldSetNode(number, title, guidance, mode, columns, result, comment);
        }

        /** Plan / asset title and description system questions are skipped when the document shows them anyway. */
        private boolean isShownElsewhere(FieldModel field) {
            if (!(field.getData() instanceof SystemFieldDataModel system) || system.getSystemFieldType() == null) return false;
            return switch (system.getSystemFieldType()) {
                case DescriptionTitle, DescriptionDescription -> true;
                case PlanTitle, PlanDescription -> this.options.isCover();
                default -> false;
            };
        }

        private PictureRenderData logo(FileEnvelopeModel logo) {
            if (logo == null) return null;
            try {
                byte[] bytes = properties.isUseSharedStorage() && !isBlank(logo.getFileRef()) ? fileStorageService.readFile(logo.getFileRef()) : logo.getFile();
                if (bytes == null || bytes.length == 0) return null;
                PictureType type = "image/png".equals(logo.getMimeType()) ? PictureType.PNG : "image/jpeg".equals(logo.getMimeType()) ? PictureType.JPEG : null;
                if (type == null) return null;
                try (ImageInputStream input = ImageIO.createImageInputStream(new ByteArrayInputStream(bytes))) {
                    Iterator<ImageReader> readers = ImageIO.getImageReaders(input);
                    if (!readers.hasNext()) return null;
                    ImageReader reader = readers.next();
                    reader.setInput(input);
                    double width = reader.getWidth(0);
                    double height = reader.getHeight(0);
                    double scale = Math.min(1.0, Math.min(LOGO_MAX_WIDTH_PX / width, LOGO_MAX_HEIGHT_PX / height));
                    return Pictures.ofBytes(bytes, type).size((int) Math.round(width * scale), (int) Math.round(height * scale)).create();
                }
            } catch (Exception e) {
                logger.error("could not read the cover logo", e);
                return null;
            }
        }

        private String languageName(String code) {
            if (isBlank(code)) return null;
            Language language = languageService.getLanguage(code);
            return language != null && !isBlank(language.getName()) ? language.getName() : code;
        }

        private String date(Instant instant) {
            return instant != null ? this.dateFormatter.format(instant) : null;
        }
    }

    /** Gathers the flat question list, answers by semantic and references while a body is built. */
    private static final class Collector {
        private final String section;
        final List<QuestionView> questions;
        final Map<String, String> semantics;
        final Map<String, String> semanticLabels;
        final Map<String, List<ReferenceView>> references;

        Collector(String section) {
            this(section, new ArrayList<>(), new LinkedHashMap<>(), new LinkedHashMap<>(), new LinkedHashMap<>());
        }

        private Collector(String section, List<QuestionView> questions, Map<String, String> semantics, Map<String, String> semanticLabels, Map<String, List<ReferenceView>> references) {
            this.section = section;
            this.questions = questions;
            this.semantics = semantics;
            this.semanticLabels = semanticLabels;
            this.references = references;
        }

        Collector within(String section) {
            return new Collector(section != null ? section : this.section, this.questions, this.semantics, this.semanticLabels, this.references);
        }

        /** Takes over everything another collector gathered. */
        void absorb(Collector other) {
            this.questions.addAll(other.questions);
            other.semantics.forEach((k, v) -> this.semantics.merge(k, v, (a, b) -> a + "; " + b));
            other.semanticLabels.forEach(this.semanticLabels::putIfAbsent);
            other.references.forEach((k, v) -> this.references.computeIfAbsent(k, x -> new ArrayList<>()).addAll(v));
        }

        /** A facts row for {@code sem:<semantic>}: the answer, labelled like the question carrying the semantic. */
        FactView semanticFact(String semantic) {
            String key = Views.key(semantic);
            String value = this.semantics.get(key);
            return value == null ? null : new FactView(this.semanticLabels.getOrDefault(key, Views.humanize(semantic)), value, null);
        }

        void add(FieldModel field, String fieldSetTitle, Body.Answer answer) {
            if (answer == null || answer instanceof Body.Missing) return;
            QuestionView question = new QuestionView();
            question.label = Objects.requireNonNullElse(AnswerResolver.label(field), nonNull(fieldSetTitle));
            question.section = this.section;
            question.answer = Views.plainText(answer);
            if (field.getSemantics() != null) question.semantics.addAll(field.getSemantics());
            this.questions.add(question);
            for (String semantic : question.semantics) {
                String key = Views.key(semantic);
                if (key == null || key.isEmpty()) continue;
                this.semantics.merge(key, question.answer, (a, b) -> a + "; " + b);
                this.semanticLabels.putIfAbsent(key, question.label);
            }
            if (answer instanceof Body.References references) {
                for (ReferenceView reference : references.references()) {
                    if (reference.typeCode != null) this.references.computeIfAbsent(Views.key(reference.typeCode), k -> new ArrayList<>()).add(reference);
                }
            }
        }
    }

    /** Visibility states computed by OpenCDMP; an id without a state is hidden unless the caller says otherwise. */
    private static final class Visibility {
        private final Map<String, Boolean> states = new HashMap<>();

        Visibility(List<VisibilityStateModel> states) {
            if (states == null) return;
            for (VisibilityStateModel state : states) {
                if (state.getFieldId() != null) this.states.put(state.getFieldId() + "#" + state.getOrdinal(), Boolean.TRUE.equals(state.getVisible()));
            }
        }

        boolean isVisible(String id, Integer ordinal, boolean whenUnknown) {
            Boolean state = this.states.get(id + "#" + ordinal);
            return state != null ? state : whenUnknown;
        }
    }

    // ------------------------------------------------------------------ helpers

    private static List<RootSectionModel> sortedSections(PlanModel plan) {
        if (plan.getPlanBlueprint() == null || plan.getPlanBlueprint().getDefinition() == null || plan.getPlanBlueprint().getDefinition().getSections() == null) return List.of();
        return plan.getPlanBlueprint().getDefinition().getSections().stream()
                .sorted(Comparator.comparing(SectionModel::getOrdinal, Comparator.nullsLast(Comparator.naturalOrder()))).toList();
    }

    private static ChapterView chapter(SectionView section) {
        ChapterView chapter = new ChapterView();
        chapter.section = section;
        return chapter;
    }

    private static ChapterView chapter(AssetView asset) {
        ChapterView chapter = new ChapterView();
        chapter.asset = asset;
        return chapter;
    }

    private static ChapterView group(String title, String number) {
        ChapterView chapter = new ChapterView();
        chapter.group = new GroupView();
        chapter.group.title = title;
        chapter.group.number = number;
        chapter.group.numberPrefix = number != null ? number + "\u2002" : null;
        chapter.group.heading = number != null ? number + "\u2002" + title : title;
        return chapter;
    }

    private static String childNumber(String prefix, int index) {
        return prefix == null ? String.valueOf(index) : prefix + "." + index;
    }

    private static RichText richText(String html) {
        return Views.isBlankHtml(html) ? null : new RichText(html);
    }

    private static String nonNull(String value) {
        return value == null ? "" : value;
    }

    private static boolean isBlank(String value) {
        return value == null || value.isBlank();
    }
}
