# File Transformer DOCX for OpenCDMP

**file-transformer-docx** is an implementation of the `file-transformer-base` package designed to handle the export of OpenCDMP data into **DOCX** and **PDF** formats. This service is a microservice built using **Spring Boot** and can be easily integrated with the OpenCDMP platform as an export option.

### Important Note
- **Exports**: Supported for **DOCX** and **PDF** formats.
- **Imports**: Not supported for these formats.

## Overview

This microservice allows users to export plans and descriptions from OpenCDMP into DOCX and PDF formats. It leverages the base interfaces provided by `file-transformer-base` and focuses specifically on export functionality. The exported documents are formatted in a human-readable way, making them ideal for reports and document sharing.

## Features

- **DOCX Export**: Export OpenCDMP plans and descriptions to DOCX format.
- **PDF Export**: Export OpenCDMP plans and descriptions to PDF format.
- **Spring Boot Microservice**: Built as a Spring Boot microservice for seamless integration with OpenCDMP.
- **Flexible Configuration**: Easily configurable to support various document templates.

## Key Endpoints

This service implements the following endpoints as per `FileTransformerController`:

### Export Endpoints

- **POST `/export/plan`**: Export a plan to DOCX or PDF.
- **POST `/export/description`**: Export a description to DOCX or PDF.

```bash
POST /export/plan
{
    "planModel": { ... },
    "format": "docx" // or "pdf"
}
```

```bash
POST /export/description
{
    "descriptionModel": { ... },
    "format": "docx" // or "pdf"
}
```

### Configuration Endpoint

- **GET `/formats`**: Returns supported formats for export (DOCX and PDF).

## Example

To export a plan into DOCX format:

```bash
POST /export/plan
{
    "planModel": { ... },
    "format": "docx"
}
```

To export a description into PDF format:

```bash
POST /export/description
{
    "descriptionModel": { ... },
    "format": "pdf"
}
```

## License

This repository is licensed under the [EUPL 1.2 License](LICENSE).

## Contact

For questions or support regarding this implementation, please contact:

- **Email**: opencdmp at cite.gr
