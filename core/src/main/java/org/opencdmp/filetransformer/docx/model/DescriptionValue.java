package org.opencdmp.filetransformer.docx.model;

import java.util.UUID;

public class DescriptionValue {

    private String value;
    private UUID referenceId;

    public DescriptionValue() {
    }

    public DescriptionValue(String value) {
        this.value = value;
    }

    public DescriptionValue(String value, UUID referenceId) {
        this.value = value;
        this.referenceId = referenceId;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public UUID getReferenceId() {
        return referenceId;
    }

    public void setReferenceId(UUID referenceId) {
        this.referenceId = referenceId;
    }
}
