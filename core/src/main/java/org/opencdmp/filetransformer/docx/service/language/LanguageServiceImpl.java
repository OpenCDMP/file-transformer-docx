package org.opencdmp.filetransformer.docx.service.language;

import com.fasterxml.jackson.databind.ObjectMapper;
import gr.cite.tools.logging.LoggerService;
import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.service.wordfiletransformer.WordFileTransformerServiceProperties;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.util.Arrays;
import java.util.List;

@Component
public class LanguageServiceImpl implements LanguageService {
    private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(LanguageServiceImpl.class));
    private final WordFileTransformerServiceProperties properties;
    private final ObjectMapper objectMapper = new ObjectMapper();
    private List<Language> languages = null;
    public LanguageServiceImpl(WordFileTransformerServiceProperties properties) {
	    this.properties = properties;
    }

    @Override
    public List<Language> loadLanguage() {

        if (languages != null) return languages;
        try {
            languages = Arrays.asList(objectMapper.readValue(ResourceUtils.getFile(this.properties.getLanguages()), Language[].class));
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
