package org.opencdmp.filetransformer.docx.service.language;

import com.fasterxml.jackson.databind.ObjectMapper;
import gr.cite.tools.logging.LoggerService;
import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Component;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Component
public class LanguageServiceImpl implements LanguageService {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(LanguageServiceImpl.class));
    private final WordFileTransformerServiceProperties properties;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private List<Language> languages = null;
    private final ResourceLoader resourceLoader;

    public LanguageServiceImpl(WordFileTransformerServiceProperties properties, ResourceLoader resourceLoader) {
	    this.properties = properties;
        this.resourceLoader = resourceLoader;
    }

    @Override
    public List<Language> loadLanguage() {

        if (languages != null) return languages;
        try {
            Resource resource = this.resourceLoader.getResource(this.properties.getLanguages());
            if(!resource.isReadable()) return new ArrayList<>();
            try(InputStream inputStream = resource.getInputStream()) {
                languages = Arrays.asList(objectMapper.readValue(inputStream, Language[].class));
            }
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
        }
        return languages;
    }
    @Override
    public Language getLanguage(String code){
        return this.loadLanguage().stream().filter(lng -> lng.getCode().equals(code)).findFirst().orElse(null);
    }

    protected static class LanguagesWrapper {
        private List<Language> languages;

        public List<Language> getLanguages() {
            return languages;
        }

        public void setLanguages(List<Language> languages) {
            this.languages = languages;
        }
    }
}
