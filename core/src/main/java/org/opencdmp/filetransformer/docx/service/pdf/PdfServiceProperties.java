package org.opencdmp.filetransformer.docx.service.pdf;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.ConstructorBinding;

@ConfigurationProperties(prefix = "pdf.converter")
public class PdfServiceProperties {
    private String url;
    private int maxInMemorySizeInBytes;

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public int getMaxInMemorySizeInBytes() {
        return maxInMemorySizeInBytes;
    }

    public void setMaxInMemorySizeInBytes(int maxInMemorySizeInBytes) {
        this.maxInMemorySizeInBytes = maxInMemorySizeInBytes;
    }
}
