# Writing DOCX templates

The DOCX/PDF export is driven by a Word template. OpenCDMP ships a default template for plans and one for assets. Administrators can replace either one with their own, per tenant, per plan blueprint, or per asset template. The most specific one wins.

A template is an ordinary `.docx` file with **`{{...}}` tags** in it. You design the page in Word (fonts, colours, cover page, headers and footers, where things go), and the tags mark where the plan's content goes. The engine is [poi-tl](https://github.com/Sayi/poi-tl), used in its logic-less mode. Templates can loop and show or hide blocks, but they cannot run code.

> The placeholders of the original engine (`'{OPENCDMP.PLAN.TITLE}'` and the like) are no longer supported. A template that uses them is rejected when uploaded, and exports that would use it fall back to the built-in template. To bring such a template over, start from the built-in template and carry your page design over.

## Tags

| Tag | Meaning |
| --- | --- |
| `{{plan.title}}` | Text. Replaced by the value, keeping the tag's formatting. |
| `{{?chapters}} … {{/chapters}}` | Block. Repeated once per item of a list, shown once for a non-empty value, removed for an empty value or `false`. Inside the block, names refer to the item: `{{title}}`. |
| `{{>plan.description}}` | Content. Inserts formatted content: rich text, question bodies, reference lists, tables. Put it **alone in its paragraph**; that paragraph's style is used for rich text. |
| `{{@branding.logo}}` | Picture (the configured cover logo). |

`plan`, `layout`, `labels`, `export`, `branding`, `toc` and `pageBreak` can be used anywhere, including inside blocks. Tags work in the body, in tables, and in headers and footers.

When you upload a template, it is checked. Unknown names, a plan template that never renders `chapters`, and similar mistakes are reported as warnings.

## What a template can use

### Plan exports

| Name | Contents |
| --- | --- |
| `plan.title`, `plan.kind`, `plan.version`, `plan.status`, `plan.blueprint`, `plan.language`, `plan.access`, `plan.doi`, `plan.creator` | Plan metadata as text. `kind` is the blueprint type, e.g. "Data Management Plan". |
| `plan.createdAt`, `plan.updatedAt`, `plan.finalizedAt` | Dates in the configured date format. |
| `plan.finalized` | `true` for a finalized plan: `{{?plan.finalized}}…{{/plan.finalized}}`. |
| `plan.longTitle`, `plan.shortTitle` | Whether the title is longer than about 38 characters, to set long titles smaller on a cover: `{{?plan.longTitle}}…{{/plan.longTitle}}`. `asset` has the same pair. |
| `{{>plan.description}}` | The plan description (rich text). |
| `{{>plan.facts}}` | Cover-page key facts as a two-column table; which rows, see **Cover facts** below. |
| `plan.contacts`, `plan.users` | Lists with `name`, `email`, `role`. |
| `plan.refs.<type code>` | The plan's references of a type, e.g. `{{?plan.refs.grants}}{{label}}{{/plan.refs.grants}}`, or rendered as a block with `{{>plan.refs.grants}}`. |
| `plan.sem.<semantic>` | The answer to the question carrying a semantic, dots replaced by underscores: `{{plan.sem.rda_dmp_ethical_issues_description}}`. |
| `plan.questions` | Every exported blueprint question: `label`, `answer`, `section`, `semantics`. |
| `chapters` | The document body, in order: see below. |
| `assets` | Every exported asset, in document order (e.g. for an overview table). |

Each item of `chapters` is exactly one of the following:

- **`section`**: a blueprint section. Fields: `number`, `numberPrefix` (the number plus a space, empty when numbering is off), `title`, `heading` (number and title), `{{>description}}` (its guidance, when exported), `newPage`, `{{>body}}` (its questions), `assets`, `questions`.
- **`asset`**: an asset. Fields: `number`, `numberPrefix`, `title`, `heading` (number and title), `meta` (type · template · status), `type`, `template`, `status`, `section`, `tags`, `newPage`, `{{>description}}`, `{{>body}}`, `{{>facts}}`, `refs`, `sem`, `questions`.
- **`group`**: a heading introducing grouped assets. Fields: `number`, `numberPrefix`, `title`, `heading`, `newPage`.

Assets continue the numbering of where they appear: after their section they are sub-chapters of it (section 8, asset 8.1, the asset's sections 8.1.1); grouped before or after the sections, the Assets chapter is numbered like a section (1 before them, the next number after them) and its assets are 1.1, 1.2… or 9.1, 9.2…. In an export of a single asset its sections start at 1.

`newPage` follows the page-break layout options and is never set on the first chapter or on the chapter right after a group heading, so `{{?newPage}}{{>pageBreak}}{{/newPage}}` never produces an empty page.

Where each asset appears in `chapters` follows the **Asset placement** layout option, so one template supports every placement. A minimal body looks like this:

```
{{?chapters}}
{{?group}}{{title}}{{/group}}                      ← style: Heading 1
{{?section}}
{{numberPrefix}}{{title}}                          ← style: Heading 1
{{>body}}
{{/section}}
{{?asset}}
{{numberPrefix}}{{title}}                          ← style: OpenCDMP Asset Title
{{>body}}
{{/asset}}
{{/chapters}}
```

### Asset exports

Asset exports provide `asset` (the fields listed above, with sections starting at Heading 1) and `plan` (the asset's plan). The body is `{{>asset.body}}`.

### Everywhere

| Name | Contents |
| --- | --- |
| `layout.cover`, `layout.toc`, `layout.numbering`, `layout.assetPageBreak`, `layout.assetsInSections`, `layout.assetsBeforeSections`, `layout.assetsAfterSections` | The layout options, as `true` / `false`, for `{{?layout.cover}}…{{/layout.cover}}`. |
| `labels.<key>` | Interface texts in the plan's language (`version`, `contents`, `page`, `of`, `partOf`, …; see `core/src/main/resources/export-labels/`). |
| `export.date`, `export.generator`, `export.note` | The export date, "Generated with OpenCDMP" (absent when blanked through **Wording**), and both as one line. |
| `{{>toc}}` | A table of contents (heading levels 1–2), already listing the chapters; the PDF conversion adds page numbers, and in Word you press F9. |
| `{{>pageBreak}}` | A page break. |
| `{{@branding.logo}}` | The cover logo, if one is configured. |

## Styles

Generated content uses named styles only, so you restyle it by editing these styles in Word (*Home → Styles → Modify*). If your template does not define one of them, the definition from the default template is copied in.

| Style | Used for |
| --- | --- |
| Heading 1–6 | Section headings (a plan section is Heading 1, its sub-sections Heading 2…; an asset's sections start at Heading 3 in a plan and at Heading 1 in an asset export). |
| OpenCDMP Field Set Title | The title of a group of questions. |
| OpenCDMP Title Long | The cover title when it is long (the default templates switch to it with `longTitle`). |
| OpenCDMP Asset Title, OpenCDMP Asset Meta, OpenCDMP Asset Lead | An asset's heading band (title and meta line share its borders and fill, so Word draws them as one box) and its description. |
| OpenCDMP Guidance | Section and question help texts, when **Question guidance** is on. |
| OpenCDMP Question / OpenCDMP Answer | A question's label / its answer. |
| OpenCDMP Entry | "Entry 2" above repeated question groups. |
| OpenCDMP Reference Details | The details line under a reference ("ORCID … · Affiliation …"). |
| OpenCDMP Placeholder / OpenCDMP Comment | "Not answered" / field-set comments. |
| OpenCDMP Table Header / OpenCDMP Table Text | Table cells. |
| OpenCDMP Fact Label / OpenCDMP Fact Value | The key-facts table. |
| Caption, TOC 1, TOC 2 | Figure captions, table of contents. |
| Hyperlink, OpenCDMP Strong Char, OpenCDMP Number Char, OpenCDMP Meta Char | Links, reference names, heading numbers, muted text. |
| OpenCDMP Table Lined / Grid / Striped, OpenCDMP Table Facts | Table styles (the **Table style** layout option picks one of the first three). |

The default templates use the accent colour `#2457C5`. When an **Accent Color** is configured, every use of it (and of its 10% tint in table headers) is replaced. Use the same colour in your own template if you want it to follow the configured accent.

## Layout options

Administrators set these in the DOCX plugin's configuration: on the tenant configuration page, and per plan blueprint and asset template, with the most specific setting winning. The editors show each setting's default ("Default (Lined)"); a setting left empty follows the broader scope and then that default, so change a setting only where it should differ.

| Option | Values |
| --- | --- |
| Asset placement | after each section's content (default) · grouped before the sections · grouped after the sections |
| Question layout | question above answer (default) · question / answer table · answers only |
| Reference layout | name and details (default) · name only · table |
| Reference details | Which reference fields to show per reference type, in order, e.g. `researchers:pid,affiliation; grants:acronym,startDate`; `pid` is the persistent identifier (ORCID, DOI, …). By default the identifier and every field of the reference type are shown. |
| Table style | lined (default) · grid · striped |
| Cover facts (plans / assets) | The rows of the key-facts table, see below. |
| Question guidance | Leave out (default) · show the help texts of sections and questions. |
| Comments | Include (default) · leave out field-set comments. |
| Section page breaks | Start each blueprint section on a new page. |
| Wording | Replace interface texts: `key=text` pairs separated by `;`, e.g. `asset=Dataset; planKind=Output Management Plan; generatedWith=` (an empty text removes it). Keys are those of `labels` above; `planKind` also replaces the blueprint type on the cover. |
| Draft watermark, Draft watermark text | Watermark exports that are not finalized (default), or never. The text defaults to the plan's or asset's status in capitals ("UNDER REVIEW"), or "DRAFT" in the plan's language when the status has no name; a configured text is used for every status. |
| Page size, Date format, Unanswered questions | A4 (default) or Letter; a date pattern (default `dd/MM/yyyy`); leave out unanswered questions (default) or keep them marked "Not answered". |
| Cover page, Table of contents, Section numbering, Asset page breaks | On or off. |

### Cover facts

A comma-separated list of rows, shown in that order; `key=Label` renames a row, and rows without a value are left out.

- **Plans** (default `funders, grants, researchers, organisations, contacts, doi`): any reference type code (the plan's references of that type, from every blueprint section, labelled like the blueprint question asking for them), `contacts`, `doi`, `version`, `status`, `blueprint`, `language`, `access`, `created`, `updated`, `finalized`, `creator`, `members`, or `sem:<semantic>` (the answer to the question carrying that semantic).
- **Assets** (default `type, template, plan, section, status, tags`): those keys, `created`, any reference type code (the asset's references of that type), or `sem:<semantic>`.

Example: `grants=Project, funders, licenses=Licence, sem:rda.dmp.ethical_issues_exist=Ethics`.

Users are not asked anything at export time: page size, date format and unanswered questions are organisation settings, and values users stored for them earlier are ignored. The section and asset pickers of the export dialog are switched off for now (commented out under `user-configuration-fields` in `word-file-transformer.yml`); the transformer still honours `include-sections` and `exclude-descriptions` when they are sent.

## Starting from the default template

The default templates can also be downloaded from the plugin's settings page ("Download the built-in template"). They are generated by [`tools/templates/build_templates.py`](../tools/templates/build_templates.py). To make your own, download [`plan.docx`](../web/src/main/resources/documents/plan.docx), edit it in Word, and upload it. The fonts it uses (Noto Sans) are embedded, so the result looks the same everywhere.
