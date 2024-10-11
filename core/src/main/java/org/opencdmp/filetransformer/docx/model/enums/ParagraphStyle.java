package org.opencdmp.filetransformer.docx.model.enums;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;
import org.opencdmp.commonmodels.enums.EnumUtils;
import org.opencdmp.commonmodels.enums.EnumValueProvider;

import java.util.Map;

public enum ParagraphStyle implements EnumValueProvider<Integer> {
    TEXT(0), HEADER1(1), HEADER2(2), HEADER3(3), HEADER4(4), TITLE(5), FOOTER(6), COMMENT(7), HEADER5(8), HEADER6(9), HTML(10), IMAGE(11);

    private final Integer value;

    ParagraphStyle(Integer value) {
        this.value = value;
    }

    @JsonValue
    public Integer getValue() {
        return value;
    }


    private static final Map<Integer, ParagraphStyle> map = EnumUtils.getEnumValueMap(ParagraphStyle.class);

    @JsonCreator
    public static ParagraphStyle of(Integer i) {
        return map.get(i);
    }
}
