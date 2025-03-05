package org.opencdmp.filetransformer.docx.service.language;

import org.opencdmp.filetransformer.docx.model.Language;
import org.opencdmp.filetransformer.docx.model.PidLink;

import java.io.FileNotFoundException;
import java.util.List;

public interface LanguageService {
	Language getLanguage(String code);
	List<Language> loadLanguage() throws FileNotFoundException;
}
