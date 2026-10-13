package org.opencdmp.filetransformer.docx.service.pdf;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.ConstructorBinding;

@ConfigurationProperties(prefix = "pdf.converter")
public class PdfServiceProperties {
    private String url;
    private int maxInMemorySizeInBytes;
    private String pdfaFormat;
    private boolean pdfua;

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public String getPdfaFormat() {
        return pdfaFormat;
    }

    public void setPdfaFormat(String pdfaFormat) {
        this.pdfaFormat = pdfaFormat;
    }

    public boolean isPdfua() {
        return pdfua;
    }

    public void setPdfua(boolean pdfua) {
        this.pdfua = pdfua;
    }

    public int getMaxInMemorySizeInBytes() {
        return maxInMemorySizeInBytes;
    }

    public void setMaxInMemorySizeInBytes(int maxInMemorySizeInBytes) {
        this.maxInMemorySizeInBytes = maxInMemorySizeInBytes;
    }
}
