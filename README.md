# File Transformer DOCX for OpenCDMP

`file-transformer-docx` is an implementation of the [file-transformer-base](https://github.com/OpenCDMP/file-transformer-base) package designed to handle the export of OpenCDMP data into DOCX and PDF formats. This service is a Spring Boot microservice that integrates with the OpenCDMP platform as an export option.

## Important Note
- Exports: Supported for DOCX and PDF formats.
- Imports: Not supported for these formats.

## Overview

This microservice allows users to export plans and descriptions from OpenCDMP into DOCX and PDF formats. The exported documents are formatted in a human-readable way, making them ideal for reports and document sharing.

## Features

- **DOCX Export**: Export OpenCDMP plans and descriptions to DOCX format
- **PDF Export**: Export OpenCDMP plans and descriptions to PDF format
- **Custom Templates**: Support for custom DOCX templates with field codes
- **Spring Boot Microservice**: Built for seamless integration with OpenCDMP
- **Template Hierarchy**: Blueprint/template-specific templates override tenant-level defaults

---

## API endpoints

This service implements the following endpoints as per `FileTransformerController`:

### Export endpoints

#### Export a plan

Export a plan to DOCX or PDF format.

**Endpoint**: `POST /export/plan`

**Request body**:
```json
{
  "planModel": {
    "id": "plan-uuid",
    "title": "My Research Plan",
    "description": "Plan content",
    // more data
  },
  "format": "docx"
}
```

**Supported formats**: `docx`, `pdf`

**Response**: Object that contains binary file (DOCX or PDF)

---

#### Export a description

Export a description to DOCX or PDF format.

**Endpoint**: `POST /export/description`

**Request body**:
```json
{
  "descriptionModel": {
    "id": "description-uuid",
    "title": "Dataset Description",
    "description": "Description content",
    // more data
  },
  "format": "pdf"
}
```

**Supported formats**: `docx`, `pdf`

**Response**: Object that contains binary file (DOCX or PDF)

---

## Custom templates

The File Transformer DOCX service supports custom DOCX templates with field codes for personalized export formatting.

For detailed information on:
- Creating and uploading custom templates
- Available template field codes for OpenCDMP plans and descriptions
- Template hierarchy and override behavior

See the [OpenCDMP File Transformers documentation](https://opencdmp.github.io/optional-services/file-transformers/).

---

## Integration with OpenCDMP

To integrate this service with your OpenCDMP deployment, configure the file transformer plugin in the OpenCDMP admin interface.

For detailed integration instructions, see see the [File Transformers  DOCX configuration](https://opencdmp.github.io/getting-started/configuration/backend/file-transformers/#docx-file-transformer) and the [OpenCDMP File Transformers Service Authentication](https://opencdmp.github.io/getting-started/configuration/backend/#file-transformer-service-authentication).

---

## See Also

- **File Transformers Overview**: https://opencdmp.github.io/optional-services/file-transformers
- **Developer Plugin Guide**: https://opencdmp.github.io/developers/plugins/file-transformers

---

## License

This repository is licensed under the [EUPL-1.2 License](LICENSE).

---

## Contact

For questions, support, or feedback:

- **Email**: opencdmp at cite.gr
- **GitHub Issues**: https://github.com/OpenCDMP/file-transformer-docx/issues
---

*This service is part of the OpenCDMP ecosystem. For general OpenCDMP documentation, visit [opencdmp.github.io](https://opencdmp.github.io).*