package org.opencdmp.filetransformer.docx.service.pid;

import com.fasterxml.jackson.databind.ObjectMapper;
import gr.cite.tools.logging.LoggerService;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.opencdmp.filetransformer.docx.model.PidLink;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.IOException;
import java.util.List;

@Component
public class PidServiceImpl implements PidService {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(PidServiceImpl.class));
    private final WordFileTransformerServiceProperties properties;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private List<PidLink> pidLinks = null;
    public PidServiceImpl(WordFileTransformerServiceProperties properties) {
	    this.properties = properties;
    }

    @Override
    public List<PidLink> loadPidLinks() {
        if (pidLinks != null) return pidLinks;
        try {
            pidLinks = objectMapper.readValue(ResourceUtils.getFile(this.properties.getPidTemplate()), PidLinksWrapper.class).getPidLinks();
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
        }
        return pidLinks;
    }
    @Override
    public PidLink getPid(String pidType) {
        return this.loadPidLinks().stream().filter(pl -> pl.getPid().equals(pidType)).findFirst().orElse(null);
    }

    protected static class PidLinksWrapper {
        private List<PidLink> pidLinks;

        public List<PidLink> getPidLinks() {
            return pidLinks;
        }

        public void setPidLinks(List<PidLink> pidLinks) {
            this.pidLinks = pidLinks;
        }
    }
}
