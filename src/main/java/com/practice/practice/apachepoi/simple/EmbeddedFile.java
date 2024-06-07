package com.practice.practice.apachepoi.simple;

import lombok.Getter;

@Getter
public class EmbeddedFile {
    private byte[] embeddedFileByteArray;
    private EmbeddedFileFormatEnum embeddedFileFormatEnum;
    private String embeddedFileName;
    
    public EmbeddedFile(
        byte[] embeddedFileByteArray,
        EmbeddedFileFormatEnum embeddedFileFormatEnum,
        String embeddedFileName
    ) {
        this.embeddedFileByteArray = embeddedFileByteArray;
        this.embeddedFileFormatEnum = embeddedFileFormatEnum;
        this.embeddedFileName = embeddedFileName;
    }
}
