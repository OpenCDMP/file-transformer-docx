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
import org.opencdmp.filetransformerbase.models.misc.PreprocessingDescriptionModel;
import org.opencdmp.filetransformerbase.models.misc.PreprocessingPlanModel;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.AbstractMap;
import java.util.Map;

@RestController
@RequestMapping("/api/file-transformer")
public class FileTransformerController implements org.opencdmp.filetransformerbase.interfaces.FileTransformerController {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(FileTransformerController.class));

    private final FileTransformerClient fileTransformerExecutor;

    private final AuditService auditService;

    @Autowired
    public FileTransformerController(FileTransformerClient fileTransformerExecutor, AuditService auditService) {
        this.fileTransformerExecutor = fileTransformerExecutor;
	    this.auditService = auditService;
    }

    public FileEnvelopeModel exportPlan(@RequestBody PlanModel planDepositModel, @RequestParam(value = "format",required = false)String format) throws Exception {
        logger.debug(new MapLogEntry("exportPlan " + PlanModel.class.getSimpleName()).And("planDepositModel", planDepositModel).And("format", format));

        FileEnvelopeModel model = fileTransformerExecutor.exportPlan(planDepositModel, format);

        this.auditService.track(AuditableAction.FileTransformer_ExportPlan, Map.ofEntries(
                new AbstractMap.SimpleEntry<String, Object>("planDepositModel", planDepositModel),
        new AbstractMap.SimpleEntry<String, Object>("format", format)
        ));
        return model;
    }

    public FileEnvelopeModel exportDescription(@RequestBody DescriptionModel descriptionModel, @RequestParam(value = "format",required = false)String format) throws Exception {
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

    public FileTransformerConfiguration getSupportedFormats() {
        logger.debug(new MapLogEntry("getSupportedFormats"));

        FileTransformerConfiguration model = fileTransformerExecutor.getConfiguration();

        this.auditService.track(AuditableAction.FileTransformer_GetSupportedFormats);
        
        return model;
    }
}
