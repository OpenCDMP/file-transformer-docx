# OpenCDMP DOCX file transformer

**file-transformer-docx** exports [OpenCDMP](https://github.com/OpenCDMP/OpenCDMP) plans and assets as Word (DOCX) documents, and as PDF files that a [Gotenberg](https://gotenberg.dev/) service converts from the same document. It is a Spring Boot service that implements the [file-transformer-base](https://github.com/OpenCDMP/file-transformer-base) contract; its plugin id is `docx-file-transformer`. See [File transformers](https://opencdmp.github.io/optional-services/file-transformers/) in the OpenCDMP documentation.

| Operation | Plans | Assets |
| --- | --- | --- |
| Export to DOCX | Yes | Yes |
| Export to PDF | Yes | Yes |
| Import | No | No |

## Features

- The layout of an export comes from a `.docx` template with `{{...}}` tags for loops, conditional blocks and formatted content ([poi-tl](https://github.com/Sayi/poi-tl)). The service ships a built-in template for plans and one for assets. Organisations can upload their own for the whole organisation, for one plan blueprint or for one asset template, and the most specific one applies.
- Export settings control where assets appear, the cover page and its facts, the table of contents, numbering, page size, question and reference layout, table style, the wording of the fixed texts and a draft watermark, without editing a template.
- An uploaded template is validated before it is saved, and the issues are shown under the uploader. Templates with `'{OPENCDMP.X}'` placeholders are not supported: validation reports them as an error, and an export that would use one falls back to the built-in template. See [Templates with `'{OPENCDMP.…}'` placeholders](https://opencdmp.github.io/optional-services/file-transformers/docx-templates/#templates-with-opencdmp-placeholders).
- Optional PDF/A (archival) and PDF/UA (accessible) output.
- The built-in templates embed Noto Sans, so a document looks the same in Word and in the PDF.

Template authors and administrators: see [Managing templates](https://opencdmp.github.io/optional-services/file-transformers/docx-templates/), the [template reference](https://opencdmp.github.io/optional-services/file-transformers/docx-template-reference/) and the [export settings](https://opencdmp.github.io/optional-services/file-transformers/docx-export-settings/). [docs/TEMPLATES.md](docs/TEMPLATES.md) describes the tags, the data a template can use, and the style contract.

## Requirements

- An OpenCDMP installation and its Keycloak realm. The service accepts only tokens that carry the `plugins` audience.
- A [Gotenberg](https://gotenberg.dev/) service for the PDF conversion.
- To build from source: Java 21 and Maven. The service runs on Spring Boot 4.0.

## Run with Docker

The image is [`opencdmp/file-transformer-docx`](https://hub.docker.com/r/opencdmp/file-transformer-docx) on Docker Hub. It reads `app-<PROFILE>.env` from `/config` and writes its logs to `/logs`. The OpenCDMP [Docker example deployment](https://opencdmp.github.io/deployment/install/) runs it with Gotenberg.

To run it on its own, create `config/app-docker.env`:

```properties
IDP_ISSUER_URI=https://opencdmp.example.org/auth/realms/OpenCDMP
IDP_CLAIMS_ROLES_PATH=plugins.roles
SECURITY_AUDIENCE=plugins
STORAGE_PATH=/storage
PDF_CONVERTER_URL=http://gotenberg:3000
```

and start the container:

```bash
docker run -d --name file-transformer-docx \
  -e PROFILE=docker -e SERVER_PORT=8080 -p 8080:8080 \
  -v "$PWD/config:/config" -v "$PWD/logs:/logs" \
  opencdmp/file-transformer-docx:<tag>
```

At startup the service downloads the OpenID configuration from `IDP_ISSUER_URI`, so Keycloak must be reachable at that address. `/actuator/health` reports when the service is up.

## Build from source

```bash
mvn clean package
```

The build runs the tests and produces `web/target/file-transformer-docx-web-1.0.0-SNAPSHOT.jar`.

To run the JAR, set the required variables and start it. The realm and the Gotenberg service must be reachable:

```bash
export IDP_ISSUER_URI=http://opencdmp.localhost:8180/auth/realms/OpenCDMP
export IDP_CLAIMS_ROLES_PATH=plugins.roles
export SECURITY_AUDIENCE=plugins
export STORAGE_PATH=./storage
export LOGGING_PATH=./logs
export PDF_CONVERTER_URL=http://localhost:3000
java -jar web/target/file-transformer-docx-web-1.0.0-SNAPSHOT.jar
```

Without `SERVER_PORT`, the service listens on port `8084`. [Run and test a plugin locally](https://opencdmp.github.io/developers/plugins/run-locally/) shows how to get a token, call the service, and register it with a local installation.

The `Dockerfile` builds the service from source:

```bash
docker build -t file-transformer-docx .
```

A script generates the built-in templates in `web/src/main/resources/documents/`. Read [tools/templates/README.md](tools/templates/README.md) before changing them.

## Configuration

Set these as environment variables or in `app-<PROFILE>.env`:

| Variable | Purpose |
| --- | --- |
| `IDP_ISSUER_URI`, `IDP_CLAIMS_ROLES_PATH`, `SECURITY_AUDIENCE` | Keycloak realm, roles path (`plugins.roles`) and audience (`plugins`) of the tokens the service accepts. Required. |
| `PDF_CONVERTER_URL` | Address of the Gotenberg service. Required. |
| `PDF_CONVERTER_PDFA_FORMAT`, `PDF_CONVERTER_PDFUA` | Optional PDF/A format (for example `PDF/A-2b`) and PDF/UA output. |
| `WORD_FILE_TRANSFORMER_WORD_PLAN_TEMPLATE`, `WORD_FILE_TRANSFORMER_WORD_DESCRIPTION_TEMPLATE` | Replace the built-in templates for the whole installation. |
| `WORD_FILE_TRANSFORMER_PID_LINKS` | Links written for persistent identifiers such as DOIs. |
| `STORAGE_PATH`, `WORD_FILE_TRANSFORMER_USE_SHARED_STORAGE` | Folder for temporary and shared files, and whether files are exchanged with OpenCDMP through a shared folder. |
| `WORD_FILE_TRANSFORMER_*_CODE` | Codes of the reference types the service reads, when they differ from the defaults. |
| `PROFILE`, `SERVER_PORT`, `LOGGING_*` | Profile, port and logging, as for every plugin service. |

[File transformer services: DOCX](https://opencdmp.github.io/deployment/configuration/backend/file-transformers/#docx-file-transformer) and [Settings common to all plugin services](https://opencdmp.github.io/deployment/configuration/plugin-services/#settings-common-to-all-plugin-services) list every variable with its default.

### Register with OpenCDMP

Register the service with its address and the plugin id `docx-file-transformer`: for the whole installation in the API service configuration, or for one organisation with **Add plugin** under **Settings > Plugins**. See [Registering plugin services](https://opencdmp.github.io/deployment/configuration/plugin-services/).

## API

The service implements these endpoints of `FileTransformerController`:

| Method | Path | Request body | Response |
| --- | --- | --- | --- |
| `POST` | `/api/file-transformer/export/plan?format=docx` (or `pdf`) | `PlanModel` | `FileEnvelopeModel` with the file |
| `POST` | `/api/file-transformer/export/description?format=docx` (or `pdf`) | `DescriptionModel` | `FileEnvelopeModel` with the file |
| `POST` | `/api/file-transformer/validate-template?entityType=Plan` (or `Description`) | `FileEnvelopeModel` with the `.docx` template | `TemplateValidationResultModel`: `valid` and a list of `issues` |
| `GET` | `/api/file-transformer/formats` | - | `FileTransformerConfiguration` |

The import and preprocessing endpoints return an error. The models come from [common-models](https://github.com/OpenCDMP/common-models) and file-transformer-base; see the [file transformer plugin guide](https://opencdmp.github.io/developers/plugins/file-transformers/).

## Related resources

- [File transformers overview](https://opencdmp.github.io/optional-services/file-transformers/)
- [DOCX: managing templates](https://opencdmp.github.io/optional-services/file-transformers/docx-templates/), [template reference](https://opencdmp.github.io/optional-services/file-transformers/docx-template-reference/) and [export settings](https://opencdmp.github.io/optional-services/file-transformers/docx-export-settings/)
- [File transformer services configuration](https://opencdmp.github.io/deployment/configuration/backend/file-transformers/#docx-file-transformer)
- [Registering plugin services](https://opencdmp.github.io/deployment/configuration/plugin-services/)
- [File transformer plugin guide](https://opencdmp.github.io/developers/plugins/file-transformers/)

## Contributing

Report bugs and suggest changes in [GitHub Issues](https://github.com/OpenCDMP/file-transformer-docx/issues). See [Contributing](https://opencdmp.github.io/developers/contributing/) for how to propose a change.

## License

file-transformer-docx is licensed under the [European Union Public Licence (EUPL) v1.2](LICENSE). The Noto Sans fonts in `tools/templates/fonts/` are licensed under the [SIL Open Font License](tools/templates/fonts/OFL.txt).

## Contact

Questions: email opencdmp at cite.gr.
