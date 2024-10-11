package org.opencdmp.filetransformer.docx.model.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import org.opencdmp.commonmodels.enums.EnumUtils;
import org.opencdmp.commonmodels.enums.EnumValueProvider;

import java.util.Map;

public enum FileFormats implements EnumValueProvider<String> {
    DOCX("docx"),
    PDF("pdf");

    private final String value;

    FileFormats(String value) {
        this.value = value;
    }

    @JsonValue
    public String getValue() {
        return value;
    }

    private static final Map<String, FileFormats> map = EnumUtils.getEnumValueMap(FileFormats.class);

    @JsonCreator
    public static FileFormats of(String i) {
        return map.get(i);
    }
}
