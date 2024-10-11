package org.opencdmp.filetransformer.docx.service.pdf;

import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableConfigurationProperties({PdfServiceProperties.class})
public class PdfServiceConfiguration {
}
