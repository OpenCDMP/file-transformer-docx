package org.opencdmp.filetransformer.docx.service.pid;

import com.fasterxml.jackson.databind.ObjectMapper;
import gr.cite.tools.logging.LoggerService;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.opencdmp.filetransformer.docx.model.PidLink;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Component
public class PidServiceImpl implements PidService {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(PidServiceImpl.class));
    private final WordFileTransformerServiceProperties properties;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private List<PidLink> pidLinks = null;
    private final ResourceLoader resourceLoader;

    public PidServiceImpl(WordFileTransformerServiceProperties properties, ResourceLoader resourceLoader) {
	    this.properties = properties;
        this.resourceLoader = resourceLoader;
    }

    @Override
    public List<PidLink> loadPidLinks() {
        if (pidLinks != null) return pidLinks;
        try {
            Resource resource = this.resourceLoader.getResource(this.properties.getPidTemplate());
            if(!resource.isReadable()) return new ArrayList<>();
            try(InputStream inputStream = resource.getInputStream()) {
                pidLinks = objectMapper.readValue(inputStream, PidLinksWrapper.class).getPidLinks();
            }
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
