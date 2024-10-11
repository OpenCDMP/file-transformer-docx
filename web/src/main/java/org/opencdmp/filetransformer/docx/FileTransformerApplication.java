package org.opencdmp.filetransformer.docx;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication(scanBasePackages = {
        "org.opencdmp.filetransformerbase",
        "org.opencdmp.filetransformer.docx.*",
        "gr.cite.tools",
        "gr.cite.commons"
})
public class FileTransformerApplication {

    public static void main(String[] args) {
        SpringApplication.run(FileTransformerApplication.class, args);
    }
}
