package com.practice.practice.apachepoi.simple;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

/**
 * WorkBook과 Sheet를 조정할 수 있는 컨트롤러.
 * 작업을 종료하면 close() 메소드로 Workbook 자원을 반환해야한다.
 */
public class ExcelContoller{
    private XSSFWorkbook workbook;
    private XSSFSheet workSheet;
    private XSSFRow workRow;
    private XSSFCell workCell;
    private Map<String, CellController> cellControllerMap = new HashMap<>();
    private CellController cellController;
    private Map<String, Integer> imageIndexMap = new HashMap<>(); // key: imageKey(사용자 지정), value: imageNumber(Workbook.addPicture())
    private Map<String, Integer> embeddedFileIndexMap = new HashMap<>(); // key: fileName(사용자 지정), value: embbedFileNumber(Workbook.addOlePackage())
    private List<XSSFSheet> sheetList = new ArrayList<>();

    /**
     * 내부적으로 Workbook과 Sheet를 생성한다.
     * 기본적인 파일 아이콘(MS-Office, Txt, Etc-File)이 WorkBook에 추가된다.
     * 작업을 종료하면 close() 메소드로 Workbook 자원을 반환해야한다.
     * @throws IOException
     */
    public ExcelContoller() throws IOException{
        this("sheet1", (short) 10);
    }

    /**
     * 내부적으로 Workbook과 Sheet를 생성한다.
     * 기본적인 파일 아이콘(MS-Office, Txt, Etc-File)이 WorkBook에 추가된다.
     * 작업을 종료하면 close() 메소드로 Workbook 자원을 반환해야한다.
     * @throws IOException
     */
    public ExcelContoller(String sheetName) throws IOException{
        this(sheetName, (short) 10);
    }

    /**
     * 내부적으로 Workbook과 Sheet를 생성한다.
     * 기본적인 파일 아이콘(MS-Office, Txt, Etc-File)이 WorkBook에 추가된다.
     * 작업을 종료하면 close() 메소드로 Workbook 자원을 반환해야한다.
     * @throws IOException
     */
    public ExcelContoller(short fontPoint) throws IOException{
        this("sheet1", fontPoint);
    }

    /**
     * 내부적으로 Workbook과 Sheet를 생성한다.
     * 기본적인 파일 아이콘(MS-Office, Txt, Etc-File)이 WorkBook에 추가된다.
     * 작업을 종료하면 close() 메소드로 Workbook 자원을 반환해야한다.
     * @throws IOException
     */
    public ExcelContoller(String sheetName, short fontPoint) throws IOException{
        workbook = new XSSFWorkbook();
        workSheet = workbook.createSheet(sheetName);
        sheetList.add(workSheet);
        registIconImage();
        workbook.getCellStyleAt(0).getFont().setFontName(UnitConverter.BASE_FONT_NAME);
        workbook.getCellStyleAt(0).getFont().setFontHeightInPoints(fontPoint);
    }

    /**
     * XSSFWorkbook을 반환한다.
     * @return XSSFWorkbook을 반환한다.
     */
    public XSSFWorkbook getWorkbook(){
        return workbook;
    }

    /**
     * List<XSSFSheet>을 반환한다.
     * @return List<XSSFSheet>을 반환한다.
     */
    public List<XSSFSheet> getSheetList(){
        return sheetList;
    }

    /**
     * index번호에 맞는 XSSFSheet를 반환한다.
     * @param index Sheet 번호(0부터 시작)
     * @return index번호에 맞는 XSSFSheet를 반환한다.
     */
    public XSSFSheet getSheet(int index){
        return sheetList.get(index);
    }

    /**
     * 현재 작업중인 XSSFSheet를 반환한다.
     * @return 현재 작업중인 XSSFSheet를 반환한다.
     */
    public XSSFSheet getActivatedSheet(){
        return workSheet;
    }

    /**
     * index번호로 작업중인 Sheet를 설정한다.
     * @param index Sheet 번호(0부터 시작)
     * @return 현재 인스턴스(ExcelContoller)
     */
    public ExcelContoller selectSheet(int index){
        workSheet = sheetList.get(index);
        workbook.setActiveSheet(index);
        return this;
    }

    /**
     * 작업중인 Sheet의 이름을 sheetName으로 설정한다.
     * @param sheetName 변경할 Sheet의 이름
     * @return 현재 인스턴스(ExcelContoller)
     */
    public ExcelContoller setSheetName(String sheetName){
        workbook.setSheetName(workbook.getActiveSheetIndex(), sheetName);
        return this;
    }

    /**
     * Column의 Width를 pixel으로 변경한다.
     * Excel의 width 특성상 완벽하게 pixel로 구현하기가 쉽지 않아서 오차가 있다.
     * @param columnIndex 변경할 Column의 번호(0부터 시작).
     * @param pixel
     * @return 현재 인스턴스(ExcelContoller)
     */
    public ExcelContoller setColumnWidthInPixel(int columnIndex, int pixel){
        double fontPoint = workbook.getCellStyleAt(0).getFont().getFontHeightInPoints();
        workSheet.setColumnWidth(columnIndex, UnitConverter.widthPixelToWidth(pixel, fontPoint));
        return this;
    }

    /**
     * Row의 Height를 pixel으로 변경한다.
     * @param rowIndex 변경할 Row의 번호(0부터 시작).
     * @param pixel
     * @return 현재 인스턴스(ExcelController)
     */
    public ExcelContoller setRowHeightInPixel(int rowIndex, int pixel){
        XSSFRow row = workSheet.getRow(rowIndex);
        if(row == null){
            row = workSheet.createRow(rowIndex);
        }
        row.setHeight(UnitConverter.getHeightFromPixel(pixel));
        return this;
    }

    /**
     * Row의 Height를 point으로 변경한다.
     * @param rowIndex 변경할 Row의 번호(0부터 시작).
     * @param point
     * @return 현재 인스턴스(ExcelController)
     */
    public ExcelContoller setRowHeight(int rowIndex, int point){
        XSSFRow row = workSheet.getRow(rowIndex);
        if(row == null){
            row = workSheet.createRow(rowIndex);
        }
        row.setHeight(UnitConverter.getHeightFromPoint(point));
        return this;
    }

    /**
     * createObjectData에 쓰일 기본 아이콘 이미지를 등록한다.
     * icon 이미지 파일은 /static/poi에서 불러온다.
     * @throws IOException
     */
    private void registIconImage() throws IOException{
        List<String> iconFileNames = new ArrayList<>();
        for(EmbeddedFileFormatEnum fileFormatEnum : EmbeddedFileFormatEnum.values()){
            iconFileNames.add(fileFormatEnum.getIconName());
        }

        for(String iconFileName : iconFileNames){
            byte[] imageByteArray = Files.readAllBytes(new ClassPathResource("static/poi/" + iconFileName).getFile().toPath());
            int imageIndex = workbook.addPicture(imageByteArray, ImageFormatEnum.PICTURE_TYPE_PNG.getValue());
            imageIndexMap.put(iconFileName, imageIndex);
        }
    }

    /**
     * cellControllerMap에 쓰이는 Key를 생성한다.
     * @param rowIndex Row의 번호(0부터 시작).
     * @param colInex Column의 번호(0부터 시작).
     * @return "R{rowIndex}C{colIndex}"
     */
    private String getCellControllerKey(int rowIndex, int colInex){
        return "R" + rowIndex + "C" + colInex;
    }

    /**
     * 작업할 Cell을 선택한다.
     * @param rowIndex Row의 번호(0부터 시작).
     * @param colIndex Column의 번호(0부터 시작).
     * @return Cell을 조정할 수 있는 CellController 인스턴스를 반환한다.
     */
    public CellController selectCell(int rowIndex, int colIndex){
        String cellControllerKey = getCellControllerKey(rowIndex, colIndex);

        if(cellControllerMap.containsKey(cellControllerKey)){
            return cellControllerMap.get(cellControllerKey);
        }else{
            workRow = workSheet.getRow(rowIndex);
            if(workRow == null){
                workRow = workSheet.createRow(rowIndex);
            }
            workCell = workRow.createCell(colIndex);
            cellController = new CellController(workCell, imageIndexMap, embeddedFileIndexMap);
            cellControllerMap.put(cellControllerKey, cellController);
            return cellController;
        }
    }

    /**
     * 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @return 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @throws IOException
     */
    public ByteArrayOutputStream getByteArrayOutputStream() throws IOException{
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        workbook.write(result);
        return result;
    }

    /**
     * 작업한 Workbook을 byte[]으로 반환한다.
     * @return 작업한 Workbook을 byte[]으로 반환한다.
     * @throws IOException
     */
    public byte[] getByteArray() throws IOException{
        ByteArrayOutputStream byteArrayOutputStream = getByteArrayOutputStream();
        return byteArrayOutputStream.toByteArray();
    }

    /**
     * 작업한 Workbook을 Resource으로 반환한다.
     * @return 작업한 Workbook을 Resource으로 반환한다.
     * @throws IOException
     */
    public Resource getResource() throws IOException{
        ByteArrayOutputStream byteArrayOutputStream = getByteArrayOutputStream();
        return new ByteArrayResource(byteArrayOutputStream.toByteArray());
    }

    /**
     * Workbook을 close한다.
     * @throws IOException
     */
    public void close() throws IOException{
        workbook.close();
    }

    /**
     * 작업한 Workbook을 ByteArrayOutputStream으로 반환하고,
     * Workbook을 close한다.
     * @return 작업한 Workbook을 ByteArrayOutputStream으로 반환한다.
     * @throws IOException
     */
    public ByteArrayOutputStream getByteArrayOutputStreamAndClose() throws IOException{
        ByteArrayOutputStream result = new ByteArrayOutputStream();
        workbook.write(result);
        workbook.close();
        return result;
    }

    /**
     * 작업한 Workbook을 byte[]으로 반환하고,
     * Workbook을 close한다.
     * @return 작업한 Workbook을 byte[]으로 반환한다.
     * @throws IOException
     */
    public byte[] getByteArrayAndClose() throws IOException{
        ByteArrayOutputStream byteArrayOutputStream = getByteArrayOutputStream();
        workbook.close();
        return byteArrayOutputStream.toByteArray();
    }

    /**
     * 작업한 Workbook을 Resource으로 반환하고,
     * Workbook을 close한다.
     * @return 작업한 Workbook을 Resource으로 반환한다.
     * @throws IOException
     */
    public Resource getResourceAndClose() throws IOException{
        ByteArrayOutputStream byteArrayOutputStream = getByteArrayOutputStream();
        workbook.close();
        return new ByteArrayResource(byteArrayOutputStream.toByteArray());
    }

    /**
     * Cell을 Merge한다.
     * @param startRowIndex 시작 Row Index(0부터 시작)
     * @param endRowIndex 종료 Row Index(0부터 시작)
     * @param startColIndex 시작 Col Index(0부터 시작)
     * @param endColInex 종료 Col Index(0부터 시작)
     * @return 현재 인스턴스(ExcelController)
     */
    public ExcelContoller mergedRegion(int startRowIndex, int endRowIndex, int startColIndex, int endColInex){
        workSheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColInex));
        return this;
    }

    /**
     * Cell을 Merge한다.
     * @param startRowIndex 시작 Row Index(0부터 시작)
     * @param endRowIndex 종료 Row Index(0부터 시작)
     * @param startColIndex 시작 Col Index(0부터 시작)
     * @param endColInex 종료 Col Index(0부터 시작)
     * @return Mearge된 영역의 Cell을 조정할 수 있는 CellController 인스턴스를 반환한다.
     */
    public CellController mergedRegionAndSelect(int startRowIndex, int endRowIndex, int startColIndex, int endColInex){
        workSheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColInex));
        return selectCell(startRowIndex, startColIndex);
    }

    /**
     * Cell을 Merge한다.
     * @param startCell ex) "A1"
     * @param endCell ex) "C1"
     * @return 현재 인스턴스(ExcelController)
     */
    public ExcelContoller mergedRegion(String startCell, String endCell){
        workSheet.addMergedRegion(CellRangeAddress.valueOf(startCell + ":" + endCell));
        return this;
    }
}
