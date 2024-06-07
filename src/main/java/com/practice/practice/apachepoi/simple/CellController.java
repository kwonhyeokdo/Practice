package com.practice.practice.apachepoi.simple;

import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.awt.Color;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFObjectData;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDvAspect;

/**
 * 셀을 조정할 수 있는 컨트롤러.
 */
public class CellController {
    private XSSFWorkbook workbook;
    private XSSFSheet workSheet;
    private XSSFCell workCell;
    private XSSFCellStyle workCellStyle;
    private XSSFRow workRow;
    private XSSFDrawing workDrawing;
    private XSSFPicture workPicture;
    private XSSFFont workFont;

    private Map<String, Integer> embeddedFileIndexMap;
    private Map<String, Integer> imageIndexMap;

    /**
     * CellStyle의 기본값을 세팅한다.(글자 위쪽 맞춤, 텍스트 줄 바꿈)
     * @param workCell
     * @param imageIndexMap
     */
    protected CellController(
        XSSFCell workCell,
        Map<String, Integer> imageIndexMap,
        Map<String, Integer> embeddedFileIndexMap
    ){
        this.workCell = workCell;
        workSheet = workCell.getSheet();
        workbook = workSheet.getWorkbook();
        workCellStyle = workbook.createCellStyle();
        workRow = workCell.getRow();
        workDrawing = workSheet.createDrawingPatriarch();
        workCellStyle.setVerticalAlignment(VerticalAlignment.TOP); // 글자 위쪽 맞춤
        workCellStyle.setWrapText(true); // 텍스트 줄 바꿈
        workCell.setCellStyle(workCellStyle);

        this.imageIndexMap = imageIndexMap;
        this.embeddedFileIndexMap = embeddedFileIndexMap;
    }

    /**
     * 내용의 수직 정렬을 설정한다.
     * @param verticalAlignment ENUM
     * @return 현재 인스턴스(CellController)
     */
    public CellController setVerticalAlignment(VerticalAlignment verticalAlignment){
        workCellStyle.setVerticalAlignment(verticalAlignment);
        return this;
    }

    /**
     * 내용의 수평 정렬을 설정한다.
     * @param horizontalAlignment ENUM
     * @return 현재 인스턴스(CellController)
     */
    public CellController setHorizontalAlignment(HorizontalAlignment horizontalAlignment){
        workCellStyle.setAlignment(horizontalAlignment);
        return this;
    }

    /**
     * Cell(Column)의 Width를 pixel으로 변경한다.
     * Excel의 width 특성상 완벽하게 pixel로 구현하기가 쉽지 않아서 오차가 있다.
     * @param pixel 픽셀
     * @return 현재 인스턴스(CellController)
     */
    public CellController setWidthInPixel(int pixel){
        int colIndex = workCell.getColumnIndex();
        double fontPoint = workbook.getCellStyleAt(0).getFont().getFontHeightInPoints();
        workSheet.setColumnWidth(colIndex, UnitConverter.widthPixelToWidth(pixel, fontPoint));
        return this;
    }

    /**
     * Cell(Row)의 Height를 pixel으로 변경한다.
     * @param rowIndex 변경할 Row의 번호(0부터 시작).
     * @param pixel
     * @return 현재 인스턴스(CellController)
     */
    public CellController setHeightInPixel(int pixel){
        workRow.setHeight(UnitConverter.getHeightFromPixel(pixel));
        return this;
    }

    /**
     * Cell(Row)의 Height를 point으로 변경한다.
     * @param point
     * @return 현재 인스턴스(CellController)
     */
    public CellController setHeight(int point){
        workRow.setHeight(UnitConverter.getHeightFromPoint(point));
        return this;
    }

    /**
     * Cell의 Text를 변경한다.
     * @param text 입력할 Text
     * @return 현재 인스턴스(CellController)
     */
    public CellController setText(String text){
        workCell.setCellValue(text);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return 현재 인스턴스(CellController)
     */
    public CellController setNumber(int value){
        workCell.setCellValue(value);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return 현재 인스턴스(CellController)
     */
    public CellController setNumber(float value){
        workCell.setCellValue(value);
        return this;
    }

    /**
     * Cell의 Number를 변경한다.
     * @param value 입력할 value
     * @return 현재 인스턴스(CellController)
     */
    public CellController setNumber(double value){
        workCell.setCellValue(value);
        return this;
    }

    /**
     * Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     * @return Image 처리를 위해 사용되는 등록된 ImageKey들을 Set형태로 리턴한다.
     */
    public Set<String> getImageKeySet(){
        return imageIndexMap.keySet();
    }

    /**
     **<pre>
     **1. Cell에 Image를 넣는다.
     **2. imageObject의 imageKey가 기존에 사용/등록 되었다면, 기존 Image byte[]를 사용한다. 따라서 기존에 등록된 imageKey를 넣으면 기존 이미지를 사용할 수 있고, 새로운 Image를 사용하기 위해서는 imageKey 값이 중복되지 않게 설정해야 한다.
     **3. 기존의 등록된 ImageKey getImageKeySet()로 확인한다.
     **4. positionObject의 dx, dy의 기준은 px이다.
     **5. Cell의 Width와 Height를 초과하도록 px가 설정되어도 Image의 실제 크기는 Width와 Height 보다 클 수 없다.
     **6. 참고 - https://stackoverflow.com/questions/47503477/apache-poi-write-image-and-text-excel
     * </pre>
     * @param imageObject
     * @param positionObject
     * @return 현재 인스턴스(CellController)
     */
    public CellController setImage(ImageObject imageObject, PositionObject positionObject){
        final String imageKey = imageObject.getImageKey();
        final byte[] imageByteArray = imageObject.getImageByteArray();
        final ImageFormatEnum imageFormatEnum = imageObject.getImageFormatEnum();
        final int dx1 = positionObject.getDx1();
        final int dy1 = positionObject.getDy1();
        final int dx2 = positionObject.getDx2();
        final int dy2 = positionObject.getDy2();

        int imageIndex = -99999;
        if(imageIndexMap.containsKey(imageKey)){
            imageIndex = imageIndexMap.get(imageKey);
        }else{
            imageIndex = workbook.addPicture(imageByteArray, imageFormatEnum.getValue());
            imageIndexMap.put(imageKey, imageIndex);
        }
        int rowIndex = workCell.getRowIndex();
        int colIndex = workCell.getColumnIndex();

        XSSFClientAnchor anchor = new XSSFClientAnchor();
        anchor.setRow1(rowIndex);
        anchor.setRow2(rowIndex);
        anchor.setCol1(colIndex);
        anchor.setCol2(colIndex);
        anchor.setDx1(Units.EMU_PER_PIXEL * dx1);
        anchor.setDy1(Units.EMU_PER_PIXEL * dy1);
        anchor.setDx2(Units.EMU_PER_PIXEL * dx2);
        anchor.setDy2(Units.EMU_PER_PIXEL * dy2);
        anchor.setAnchorType(AnchorType.MOVE_DONT_RESIZE);
        workPicture = workDrawing.createPicture(anchor, imageIndex);
        return this;
    }

    /**
     * 한 Line에 몇 글자가 들어갈 수 있는지 판단하여 text의 총 Line 수를 구한다.
     * @param text
     * @return 한 Line에 몇 글자가 들어갈 수 있는지 판단하여 text의 총 Line 수를 구한다.
     */
    private int getLineCountFromText(CharSequence text, double maxCharacterCountInWidth){
        int newLineCnt = 0;
        double textCnt = 0d;

        for(int i = 0; i < text.length(); i++){
            char c = text.charAt(i);

            if(c == '\n' || c == '\r'){
                if(textCnt > 0d){
                    double lineCnt = textCnt / maxCharacterCountInWidth;
                    newLineCnt += (int)(Math.ceil(lineCnt));
                    textCnt = 0d;
                }else{
                    newLineCnt++;
                }
                continue;
            }else if(c == '"' || c == '\'' || c == '.' || c == ','){
                // continue;
            }else if(c == 'l' || c == 'i' || c == 'j'){
                textCnt += 0.25;
            }else if(c == '(' || c == ')' || c == '{' || c == '}' || c == '[' || c == ']' || c == '!' || c == 'f' || c == 't' || c == 'I'){
                textCnt += 0.3333;
            }
            else if(c == ' ' || c == '-' || c == '_' || c == '*' ||  Character.isDigit(c) || (c >= 'a' && c <= 'z')){
                textCnt += 0.5;
            }else if(c >= 'A' && c <= 'Z'){
                textCnt += 0.8;
            }else{
                textCnt += 1d;
            }

            if(i == text.length() - 1){
                double lineCnt = textCnt / maxCharacterCountInWidth;
                newLineCnt += (int)(Math.ceil(lineCnt));
            }
        }

        return newLineCnt;
    }

    /**
     * 글자의 높이를 Pixel로 구한다.
     * Font Point를 Pixel로 변환한 후 + a값을 더한다.
     * a 값이 무엇인지 정확하지가 않다. 맑은고딕 10pt 기준으로 4이다.
     * @return 글자의 높이를 Pixel로 구한다.
     */
    private int getFontHeightPixel(){
        final int fontPoint = workCellStyle.getFont().getFontHeightInPoints();
        return UnitConverter.CHARACTER_HEIGHT_PIXEL_MAP.get(fontPoint);
    }

    /**
     * workbook에 적용된 기본 폰트 포인트를 반환한다.
     * @return workbook에 적용된 기본 폰트 포인트를 반환한다.
     */
    private int getBaseFontPoint(){
        return workbook.getCellStyleAt(0).getFont().getFontHeightInPoints();
    }

    /**
     * text가 한줄 또는 여러줄 일 경우 높이가 몇 Pixel인지 구한다.
     * @param text
     * @return text가 한줄 또는 여러줄 일 경우 높이가 몇 Pixel인지 구한다.
     */
    private int getTextHeightPixel(final String text){
        final int fontPoint = workCellStyle.getFont().getFontHeightInPoints();
        final int cellWidth = workSheet.getColumnWidth(workCell.getColumnIndex());
        final int cellWidthPixel = UnitConverter.widthToWidthPixel(cellWidth, getBaseFontPoint());
        final int fontPixel = UnitConverter.pointToPixel(fontPoint);
        final int fontHeightPixel = getFontHeightPixel();
        final double maxCharacterCountInWidth = cellWidthPixel / fontPixel;
        
        final int lineCnt = getLineCountFromText(text, maxCharacterCountInWidth);
        final int textHeightPixel = lineCnt * fontHeightPixel;

        return textHeightPixel;
    }
    /**
     * 텍스트를 이어서 추가한다.
     * @param text
     * @return 현재 인스턴스(CellController)
     */
    public CellController addText(final String text){
        if(text != null && text.length() > 0){
            final int maxHeightPixel = UnitConverter.pointToPixel(workRow.getHeightInPoints());
        
            String fullText = workCell.getStringCellValue() + text;
            
            int textHeightPixel = getTextHeightPixel(fullText);
            
            if(maxHeightPixel < textHeightPixel){
                setHeightInPixel(textHeightPixel);
            }

            setText(fullText);
        }

        return this;
    }

    /**
     * heightPixel이 몇 Line인지 구한다.
     * @param heightPixel
     * @return heightPixel이 몇 Line인지 구한다.
     */
    private int getLineCountFromHeightPixel(final int heightPixel){
        final int fontHeightPixel = getFontHeightPixel();
        return (int)(Math.ceil((double)heightPixel / fontHeightPixel));
    }

    /**
     * 이미지를 추가한다.
     * 이미지 크기는 셀의 넓이를 넘어서지 못한다.
     * 이미지 크기는 원본 이미지의 크기를 넘어서지 못한다.
     * 이미지의 가로 세로 크기 중 더 큰 것을 기준으로 셀의 넓이에 따라 사이즈가 조정된다.
     * @param imageObject
     * @return 현재 인스턴스(CellController)
     */
    public CellController addImage(final ImageObject imageObject){
        return addImage(imageObject, 0);
    }

    /**
     * 이미지를 추가한다.
     * 이미지 크기는 셀의 넓이를 넘어서지 못한다.
     * 이미지 크기는 원본 이미지의 크기를 넘어서지 못한다.
     * 이미지의 가로 세로 크기 중 더 큰 것을 기준으로 셀의 넓이에 따라 사이즈가 조정된다.
     * @param imageObject
     * @param padding
     * @return 현재 인스턴스(CellController)
     */
    public CellController addImage(final ImageObject imageObject, final int padding){
        if(imageObject != null){
            final int cellWidth = workSheet.getColumnWidth(workCell.getColumnIndex());
            final int cellWidthPixel = UnitConverter.widthToWidthPixel(cellWidth, getBaseFontPoint());
            final int maxHeightPixel = UnitConverter.pointToPixel(workRow.getHeightInPoints());
            final int sourceImageWidth = imageObject.getWidth();
            final int sourceImageHeight = imageObject.getHeight();
            
            double scale = 1d;
            int imageWidthPixel = 0;
            int imageHeightPixel = 0;

            if(sourceImageWidth < cellWidthPixel && sourceImageHeight < cellWidthPixel){
                imageWidthPixel = sourceImageWidth;
                imageHeightPixel = sourceImageHeight;
            }else{
                if(sourceImageWidth > sourceImageHeight){
                    scale = (double)cellWidthPixel / sourceImageWidth;
                }else{
                    scale = (double)cellWidthPixel / sourceImageHeight;
                }
                imageWidthPixel = (int)(sourceImageWidth * scale);
                imageHeightPixel = (int)(sourceImageHeight * scale);
            }

            int imageLineCount = getLineCountFromHeightPixel(imageHeightPixel);

            String text = workCell.getStringCellValue();
            int fromTextHeightPixel = getTextHeightPixel(text);

            boolean isStart = text.length() == 0;
            for(int i = (isStart ? 1 : 0); i < imageLineCount + 1; i++){ // 엑셀에서 기본으로 빈 텍스트는 한 줄임.
                text += "\n";
            }

            int toTextHeightPixel = getTextHeightPixel(text);

            if(maxHeightPixel < toTextHeightPixel){
                setHeightInPixel(toTextHeightPixel);
            }

            setImage(imageObject, new PositionObject(0 + padding, fromTextHeightPixel + padding, imageWidthPixel - padding, fromTextHeightPixel + imageHeightPixel - padding));

            setText(text);
        }

        return this;
    }

    /**
     * 파일을 추가한다.
     * 아이콘의 크기는 30 x 30 Pixel이다.
     * @param embeddedFile
     * @param padding
     * @return 현재 인스턴스(CellController)
     * @throws IOException
     */
    public CellController addFile(final EmbeddedFile embeddedFile) throws IOException{
        return addFile(embeddedFile, 0);
    }

    /**
     * 파일을 추가한다.
     * 아이콘의 크기는 30 x 30 Pixel이다.
     * @param embeddedFile
     * @return 현재 인스턴스(CellController)
     * @throws IOException
     */
    public CellController addFile(final EmbeddedFile embeddedFile, final int padding) throws IOException{
        if(embeddedFile != null){
            final int maxHeightPixel = UnitConverter.pointToPixel(workRow.getHeightInPoints());
            final int size = 30;

            int imageLineCount = getLineCountFromHeightPixel(size);

            String text = workCell.getStringCellValue();
            boolean isStart =
                (text.length() == 0) ||
                (
                    text.length() >= 2 &&
                    (
                        (text.charAt(text.length() - 1) == '\n' && text.charAt(text.length() - 2) == '\n') ||
                        (text.charAt(text.length() - 1) == '\r' && text.charAt(text.length() - 2) == '\r')
                    )
                )
            ;

            int fromTextHeightPixel = getTextHeightPixel(text);

            for(int i = (isStart ? 1 : 0); i < imageLineCount + 1; i++){ // 엑셀에서 기본으로 빈 텍스트는 한 줄임.
                text += "\n";
            }

            int toTextHeightPixel = getTextHeightPixel(text);
            
            if(maxHeightPixel < toTextHeightPixel){
                setHeightInPixel(toTextHeightPixel);
            }

            setFile(embeddedFile, new PositionObject(0 + padding, fromTextHeightPixel + padding, size - padding, fromTextHeightPixel + size - padding));

            setText(text);
        }

        return this;
    }

    /**
     * Embedded File 처리를 위해 사용되는 등록된 embeddedFileName들을 Set형태로 리턴한다.
     * @return Embedded File 처리를 위해 사용되는 등록된 embeddedFileName들을 Set형태로 리턴한다.
     */
    public Set<String> getEmbeddedFileNameSet(){
        return embeddedFileIndexMap.keySet();
    }

    /**
     **<pre>
     **1. Cell에 Embedded File를 넣는다.
     **2. embeddedFile embeddedFileName 기존에 사용/등록 되었다면, 기존 embbedFileByteArray를 사용한다. 따라서 기존에 등록/사용된 embeddedFileName를 넣으면 기존 Embedded File을 사용할 수 있고, 새로운 Embedded File를 사용하기 위해서는 embeddedFileName 값이 중복되지 않게 설정해야 한다.
     **3. 기존의 등록된 embeddedFileName getEmbeddedFileNameSet()로 확인한다.
     **4. dx는 Cell의 좌상단 끝(모서리) 기준 x축 시작 오프셋이다.
     **5. dy는 Cell의 좌상단 끝(모서리) 기준 y축 시작 오프셋이다.
     **6. 파일의 icon은 30 * 30 px 기준으로 생성된다.
     * </pre>
     * @param embeddedFile
     * @param dx1 Cell의 좌상단 끝(모서리) 기준 x축 시작 오프셋이다.
     * @param dy1 Cell의 좌상단 끝(모서리) 기준 y축 시작 오프셋이다.
     * @return
     * @throws IOException
     */
    public CellController setFile(EmbeddedFile embeddedFile, int dx1, int dy1) throws IOException{
        int size = 30;
        int dx2 = dx1 + size;
        int dy2 = dy1 + size;
        return setFile(embeddedFile, new PositionObject(dx1, dy1, dx2, dy2));
    }

    /**
     **<pre>
     **1. Cell에 Embedded File를 넣는다.
     **2. embeddedFile embeddedFileName 기존에 사용/등록 되었다면, 기존 embbedFileByteArray를 사용한다. 따라서 기존에 등록/사용된 embeddedFileName를 넣으면 기존 Embedded File을 사용할 수 있고, 새로운 Embedded File를 사용하기 위해서는 embeddedFileName 값이 중복되지 않게 설정해야 한다.
     **3. 기존의 등록된 embeddedFileName getEmbeddedFileNameSet()로 확인한다.
     **4. positionObject의 dx, dy의 기준은 px이다.
     **5. Cell의 Width와 Height를 초과하도록 px가 설정되어도 파일 아이콘 Image의 실제 크기는 Width와 Height 보다 클 수 없다.
     * </pre>
     * @param embeddedFile
     * @param positionObject
     * @return
     * @throws IOException
     */
    public CellController setFile(EmbeddedFile embeddedFile, PositionObject positionObject) throws IOException{
        final String embbedFileName = embeddedFile.getEmbeddedFileName();
        final byte[] embbedFileByteArray = embeddedFile.getEmbeddedFileByteArray();
        final EmbeddedFileFormatEnum embeddedFileFormatEnum = embeddedFile.getEmbeddedFileFormatEnum();
        final int dx1 = positionObject.getDx1();
        final int dy1 = positionObject.getDy1();
        final int dx2 = positionObject.getDx2();
        final int dy2 = positionObject.getDy2();

        int embbedFileIndex = -99999;
        if(embeddedFileIndexMap.containsKey(embbedFileName)){
            embbedFileIndex = embeddedFileIndexMap.get(embbedFileName);
        }else{
            embbedFileIndex = workbook.addOlePackage(embbedFileByteArray, embbedFileName, embbedFileName, embbedFileName);
            embeddedFileIndexMap.put(embbedFileName, embbedFileIndex);
        }
        int rowIndex = workCell.getRowIndex();
        int colIndex = workCell.getColumnIndex();
        int imageIndex = imageIndexMap.get(embeddedFileFormatEnum.getIconName());

        XSSFClientAnchor anchor = new XSSFClientAnchor();
        anchor.setRow1(rowIndex);
        anchor.setRow2(rowIndex);
        anchor.setCol1(colIndex);
        anchor.setCol2(colIndex);
        anchor.setDx1(Units.EMU_PER_PIXEL * dx1);
        anchor.setDy1(Units.EMU_PER_PIXEL * dy1);
        anchor.setDx2(Units.EMU_PER_PIXEL * dx2);
        anchor.setDy2(Units.EMU_PER_PIXEL * dy2);
        
        XSSFObjectData txtObjectData = workDrawing.createObjectData(anchor, embbedFileIndex, imageIndex);
        txtObjectData.getOleObject().setDvAspect(STDvAspect.DVASPECT_ICON); // 파일 이미지를 더블클릭 했을 때, 엑셀 기능에 의해 썸네일 형식으로 전환되는 것을 방지.

        return this;
    }

    /**
     * Cell에 설정된 XSSFCellStyle 인스턴스를 반환한다.
     * @return Cell에 설정된 XSSFCellStyle 인스턴스를 반환한다.
     */
    public XSSFCellStyle getCellStyle(){
        return workCellStyle;
    }

    /**
     * Cell의 XSSFCellStyle을 교체한다.
     * @param cellStyle
     * @return 현재 인스턴스(CellController)
     */
    public CellController setCellStyle(XSSFCellStyle cellStyle){
        workCellStyle = cellStyle;
        workCell.setCellStyle(workCellStyle);
        return this;
    }

    /**
     * R, G, B로 XSSFColor를 생성후 반환한다.
     * @param R
     * @param G
     * @param B
     * @return R, G, B로 XSSFColor를 생성후 반환한다.
     */
    private XSSFColor getColor(int R, int G, int B){
        IndexedColorMap indexedColors = workbook.getStylesSource().getIndexedColors();
        return new XSSFColor(new Color(R, G, B), indexedColors);
    }

    /**
     * Cell의 색상을 변경한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setCellColor(int R, int G, int B){
        XSSFColor color = getColor(R, G, B);

        workCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        workCellStyle.setFillForegroundColor(color);

        return this;
    }

    /**
     * Cell의 Top-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return 현재 인스턴스(CellController)
     */
    public CellController setTopBorderStyle(BorderStyle borderStyle){
        workCellStyle.setBorderTop(borderStyle);

        return this;
    }

    /**
     * Cell의 Bottom-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return 현재 인스턴스(CellController)
     */
    public CellController setBottomBorderStyle(BorderStyle borderStyle){
        workCellStyle.setBorderBottom(borderStyle);

        return this;
    }

    /**
     * Cell의 Left-Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return 현재 인스턴스(CellController)
     */
    public CellController setLeftBorderStyle(BorderStyle borderStyle){
        workCellStyle.setBorderLeft(borderStyle);

        return this;
    }

    /**
     * Cell의 Right-Border를 설정한다.
     * @param borderStyle Border의 스타일
     * @return 현재 인스턴스(CellController)
     */
    public CellController setRightBorderStyle(BorderStyle borderStyle){
        workCellStyle.setBorderRight(borderStyle);

        return this;
    }

    /**
     * Cell의 Border Style를 설정한다.
     * @param borderStyle Border의 스타일
     * @return 현재 인스턴스(CellController)
     */
    public CellController setBorderStyle(BorderStyle borderStyle){
        return this
              .setTopBorderStyle(borderStyle)
              .setBottomBorderStyle(borderStyle)
              .setLeftBorderStyle(borderStyle)
              .setRightBorderStyle(borderStyle)
        ;
    }

    /**
     * Cell의 Top-Border Color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setTopBorderColor(int R, int G, int B){
        workCellStyle.setTopBorderColor(getColor(R, G, B));
        return this;
    }

    /**
     * Cell의 Bottom-Border Color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setBottomBorderColor(int R, int G, int B){
        workCellStyle.setBottomBorderColor(getColor(R, G, B));
        return this;
    }

    /**
     * Cell의 LEFT-Border Color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setLeftBorderColor(int R, int G, int B){
        workCellStyle.setLeftBorderColor(getColor(R, G, B));
        return this;
    }

    /**
     * Cell의 Right-Border Color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setRightBorderColor(int R, int G, int B){
        workCellStyle.setRightBorderColor(getColor(R, G, B));
        return this;
    }

    /**
     * Cell의 Border Color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setBorderColor(int R, int G, int B){
        return this
              .setTopBorderColor(R, G, B)
              .setBottomBorderColor(R, G, B)
              .setLeftBorderColor(R, G, B)
              .setRightBorderColor(R, G, B)
        ;
    }

    /**
     * Cell에 삽입한 Image의 테두리 색을 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setImageLineColor(int R, int G, int B){
        workPicture.setLineStyleColor(R, G, B);
        return this;
    }

    /**
     * workFont를 반환한다.
     * workFont가 없을 경우 새로 생성 후 반환한다.
     * @return workFont를 반환한다.
     */
    private XSSFFont getWorkFont(){
        if(workFont == null){
            workFont = workbook.createFont();
            workFont.setFontName(UnitConverter.BASE_FONT_NAME);
            workCellStyle.setFont(workFont);
        };

        return workFont;
    }

    /**
     * Cell의 Font Size를 설정한다.
     * @param size 폰트 크기
     * @return 현재 인스턴스(CellController)
     */
    public CellController setFontSize(int size){
        getWorkFont().setFontHeightInPoints((short)size);
        return this;
    }

    /**
     * Cell의 font color를 설정한다.
     * @param R
     * @param G
     * @param B
     * @return 현재 인스턴스(CellController)
     */
    public CellController setFontColor(int R, int G, int B){
        getWorkFont().setColor(getColor(R, G, B));
        return this;
    }

    /**
     * Cell의 DataFormat을 설정한다.
     * 참고 [표현형식 Index] - https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
     * @param dataformatIndex 표현형식 Index, 예시) 0x31 : text
     * @return 현재 인스턴스(CellController)
     */
    public CellController setDataFormat(int dataformatIndex){
        workCellStyle.setDataFormat(dataformatIndex);
        return this;
    }

    /**
     * Cell의 DataFormat을 설정한다.
     * 참고 [표현형식 Index] - https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/BuiltinFormats.html
     * @param dataformat 표현형식 Index, 예시) "#,##0"
     * @return 현재 인스턴스(CellController)
     */
    public CellController setDataFormat(String dataformat){
        workCellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(dataformat));
        return this;
    }
}
