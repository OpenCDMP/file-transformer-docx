package org.opencdmp.filetransformer.docx.web.controller;

import gr.cite.tools.auditing.AuditService;
import gr.cite.tools.logging.LoggerService;
import gr.cite.tools.logging.MapLogEntry;
import org.opencdmp.commonmodels.models.FileEnvelopeModel;
import org.opencdmp.commonmodels.models.description.DescriptionModel;
import org.opencdmp.commonmodels.models.plan.PlanModel;
import org.opencdmp.filetransformer.docx.audit.AuditableAction;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerClient;
import org.opencdmp.filetransformerbase.interfaces.FileTransformerConfiguration;
import org.opencdmp.filetransformerbase.models.misc.DescriptionImportModel;
import org.opencdmp.filetransformerbase.models.misc.PlanImportModel;
import org.opencdmp.commonmodels.enums.PluginEntityType;
import org.opencdmp.filetransformerbase.models.misc.PreprocessingDescriptionModel;
import org.opencdmp.filetransformerbase.models.misc.PreprocessingPlanModel;
import org.opencdmp.filetransformerbase.models.misc.TemplateValidationResultModel;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.AbstractMap;
import java.util.Map;

@RestController
@RequestMapping("/api/file-transformer")
public class FileTransformerController implements org.opencdmp.filetransformerbase.interfaces.FileTransformerController {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(FileTransformerController.class));

    private final FileTransformerClient fileTransformerExecutor;

    private final AuditService auditService;

    private final WordFileTransformerServiceProperties properties;

    private final ResourceLoader resourceLoader;

    @Autowired
    public FileTransformerController(FileTransformerClient fileTransformerExecutor, AuditService auditService, WordFileTransformerServiceProperties properties, ResourceLoader resourceLoader) {
        this.fileTransformerExecutor = fileTransformerExecutor;
	    this.auditService = auditService;
        this.properties = properties;
        this.resourceLoader = resourceLoader;
    }

    /**
     * The built-in template (the one used when no template is uploaded), as a starting point for a custom one.
     * Not part of the file transformer contract: the platform offers it only for plugins that answer.
     */
    @GetMapping("/default-template")
    public ResponseEntity<byte[]> getDefaultTemplate(@RequestParam(value = "entityType", required = false) PluginEntityType entityType) throws IOException {
        logger.debug(new MapLogEntry("getDefaultTemplate").And("entityType", entityType));
        boolean description = PluginEntityType.Description.equals(entityType);
        Resource resource = this.resourceLoader.getResource(description ? this.properties.getWordDescriptionTemplate() : this.properties.getWordPlanTemplate());
        byte[] data;
        try (InputStream stream = resource.getInputStream()) {
            data = stream.readAllBytes();
        }
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + (description ? "asset-template.docx" : "plan-template.docx"));
        return ResponseEntity.ok().headers(headers).body(data);
    }

    public FileEnvelopeModel exportPlan(@RequestBody PlanModel planDepositModel, @RequestParam(value = "format",required = false)String format) {
        logger.debug(new MapLogEntry("exportPlan " + PlanModel.class.getSimpleName()).And("planDepositModel", planDepositModel).And("format", format));

        FileEnvelopeModel model = fileTransformerExecutor.exportPlan(planDepositModel, format);

        this.auditService.track(AuditableAction.FileTransformer_ExportPlan, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("planDepositModel", planDepositModel),
        new AbstractMap.SimpleEntry<String, Object>("format", format)
        ));
        return model;
    }

    public FileEnvelopeModel exportDescription(@RequestBody DescriptionModel descriptionModel, @RequestParam(value = "format",required = false)String format) {
        logger.debug(new MapLogEntry("exportDescription " + DescriptionModel.class.getSimpleName()).And("descriptionModel", descriptionModel).And("format", format));

        FileEnvelopeModel model = fileTransformerExecutor.exportDescription(descriptionModel, format);

        this.auditService.track(AuditableAction.FileTransformer_ExportDescription, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("descriptionModel", descriptionModel),
                new AbstractMap.SimpleEntry<String, Object>("format", format)
        ));
        return model;
    }

    public PlanModel importFileToPlan(@RequestBody PlanImportModel planImportModel) {

        logger.debug(new MapLogEntry("importFileToPlan " + FileEnvelopeModel.class.getSimpleName()).And("fileEnvelope", planImportModel.getFile()));

        PlanModel model = fileTransformerExecutor.importPlan(planImportModel);

        this.auditService.track(AuditableAction.FileTransformer_ImportFileToPlan, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("fileEnvelope", planImportModel.getFile())
        ));
        return model;
    }

    public DescriptionModel importFileToDescription(@RequestBody DescriptionImportModel descriptionImportModel) {
        logger.debug(new MapLogEntry("importFileToDescription " + FileEnvelopeModel.class.getSimpleName()).And("fileEnvelope", descriptionImportModel.getFile()));

        DescriptionModel model = fileTransformerExecutor.importDescription(descriptionImportModel);

        this.auditService.track(AuditableAction.FileTransformer_ImportFileToDescription, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("importFileToDescription ", descriptionImportModel.getFile())
        ));
        return model;
    }

    public PreprocessingPlanModel preprocessingPlan(FileEnvelopeModel fileEnvelopeModel) {
        logger.debug(new MapLogEntry("PreprocessingPlanModel " + FileEnvelopeModel.class.getSimpleName()).And("fileEnvelope", fileEnvelopeModel));

        PreprocessingPlanModel model = fileTransformerExecutor.preprocessingPlan(fileEnvelopeModel);

        this.auditService.track(AuditableAction.FileTransformer_PreprocessingPlan, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("importFileToPlan ", fileEnvelopeModel)
        ));
        return model;
    }

    public PreprocessingDescriptionModel preprocessingDescription(FileEnvelopeModel fileEnvelopeModel) {
        logger.debug(new MapLogEntry("PreprocessingDescriptionModel " + FileEnvelopeModel.class.getSimpleName()).And("fileEnvelope", fileEnvelopeModel));

        PreprocessingDescriptionModel model = fileTransformerExecutor.preprocessingDescription(fileEnvelopeModel);

        this.auditService.track(AuditableAction.FileTransformer_PreprocessingDescription, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("importFileToDescription ", fileEnvelopeModel)
        ));
        return model;    }

    @Override
    public TemplateValidationResultModel validateTemplate(@RequestBody FileEnvelopeModel fileEnvelopeModel, @RequestParam(value = "entityType", required = false) PluginEntityType entityType) {
        logger.debug(new MapLogEntry("validateTemplate " + FileEnvelopeModel.class.getSimpleName()).And("entityType", entityType));

        TemplateValidationResultModel model = fileTransformerExecutor.validateTemplate(fileEnvelopeModel, entityType);

        this.auditService.track(AuditableAction.FileTransformer_ValidateTemplate, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("filename", fileEnvelopeModel != null && fileEnvelopeModel.getFilename() != null ? fileEnvelopeModel.getFilename() : ""),
                new AbstractMap.SimpleEntry<String, Object>("entityType", entityType != null ? entityType : "")
        ));
        return model;
    }

    public FileTransformerConfiguration getSupportedFormats() {
        logger.debug(new MapLogEntry("getSupportedFormats"));

        FileTransformerConfiguration model = fileTransformerExecutor.getConfiguration();

        this.auditService.track(AuditableAction.FileTransformer_GetSupportedFormats);
        
        return model;
    }
}
