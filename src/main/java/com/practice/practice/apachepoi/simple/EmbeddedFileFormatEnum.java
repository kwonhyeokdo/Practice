package com.practice.practice.apachepoi.simple;

public enum EmbeddedFileFormatEnum {
    EXCEL("icon_excel.png"),
    POWER_POINT("icon_power_point.png"),
    WORD("icon_word.png"),
    TEXT("icon_txt.png"),
    ETC("icon_file.png"),
    JPG("icon_image.png"),
    PNG("icon_image.png"),
    PDF("icon_pdf.png")
    ;

    private final String iconName;

    EmbeddedFileFormatEnum(String iconName) {
        this.iconName = iconName;
    }

    public String getIconName(){
        return this.iconName;
    }
}
