package com.practice.practice.apachepoi;

import java.io.IOException;
import java.nio.file.Files;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import com.practice.practice.apachepoi.simple.ExcelContoller;
import com.practice.practice.apachepoi.simple.EmbeddedFile;
import com.practice.practice.apachepoi.simple.EmbeddedFileFormatEnum;
import com.practice.practice.apachepoi.simple.ImageFormatEnum;
import com.practice.practice.apachepoi.simple.ImageObject;

@Service
public class ApachePoiService {

    public Resource test() throws IOException{
        ExcelContoller excelContoller = new ExcelContoller();

        // 헤더
        for(int i = 0; i < 100; i++){
            if(i == 0){
                for(int j = 0; j < 100; j++){
                    excelContoller
                        .selectCell(i, j)
                        .setText("Header" + (j + 1))
                        .setCellColor(100, 255, 100)
                        .setBorderStyle(BorderStyle.THIN)
                        .setBorderColor(0, 0, 0)
                        .setVerticalAlignment(VerticalAlignment.CENTER)
                        .setHorizontalAlignment(HorizontalAlignment.CENTER)
                    ;
                }
            }

            excelContoller
                .setColumnWidthInPixel(i, 200)
                // .setRowHeightInPixel(i, 200)
            ;
        }
        excelContoller.setRowHeightInPixel(0, 30);

        // 이미지
        byte[] dogImageByteArray = Files.readAllBytes(new ClassPathResource("static/poi/dog200x200.jpg").getFile().toPath());
        ImageObject dogImageObject = new ImageObject(dogImageByteArray, ImageFormatEnum.PICTURE_TYPE_JPEG, "dog.jpg");
        byte[] catImageByteArray = Files.readAllBytes(new ClassPathResource("static/poi/cat150x100.jpg").getFile().toPath());
        ImageObject catImageObject = new ImageObject(catImageByteArray, ImageFormatEnum.PICTURE_TYPE_JPEG, "cat.jpg");
        // byte[] fubaoImageByteArray = Files.readAllBytes(new ClassPathResource("static/poi/fubao560x410.jpg").getFile().toPath());
        // ImageObject fubaoImageObject = new ImageObject(fubaoImageByteArray, ImageFormatEnum.PICTURE_TYPE_JPEG, "fubao.jpg");

        // 파일
        byte[] textByteArray = Files.readAllBytes(new ClassPathResource("static/poi/test.txt").getFile().toPath());
        byte[] powerPointByteArray = Files.readAllBytes(new ClassPathResource("static/poi/test.pptx").getFile().toPath());
        byte[] excelByteArray = Files.readAllBytes(new ClassPathResource("static/poi/test.xlsx").getFile().toPath());
        EmbeddedFile textFile = new EmbeddedFile(textByteArray, EmbeddedFileFormatEnum.TEXT, "test.txt");
        EmbeddedFile powerPointFile = new EmbeddedFile(powerPointByteArray, EmbeddedFileFormatEnum.POWER_POINT, "test.pptx");
        EmbeddedFile excelFile = new EmbeddedFile(excelByteArray, EmbeddedFileFormatEnum.EXCEL, "test.xlsx");

        for(int i = 5; i <= 21; i++){
            excelContoller
                .selectCell(1, i - 5)
                .setFontSize(i)
                .addText("fontSize: " + i)
                .addFile(powerPointFile)
                .addText("위원은 탄핵 또는 금고 이상의 형의 선고에 의하지 아니하고는 파면되지 아니한다.")
                .addFile(textFile)
                .addImage(catImageObject)
                .addText("대통령은 헌법과 법률이 정하는 바에 의하여 국군을 통수한다.")
                .addFile(excelFile)
                .addText("군사재판을 관할하기 위하여 특별법원으로서 군사법원을 둘 수 있다.")
                .addImage(dogImageObject)
            ;
        }

        // 셀 머지
        // excelContoller
        //     .mergedRegion(3, 4, 0, 0)
        //     .selectCell(3, 0)
        // ;
        // excelContoller
        //     .mergedRegionAndSelect(4, 5, 0, 1)
        //     .setCellColor(100, 100, 100)
        //     .setDataFormat(0x31);
        // ;
        // excelContoller
        //     .mergedRegion("C5", "D5")
        //     .setCellColor(0, 0, 0)
        //     .setDataFormat("#,##0")
        // ;
    ;


        return excelContoller.getResourceAndClose();
    }
}