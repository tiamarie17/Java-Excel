package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import static org.excel.Utils.ConvertTxtFile;

public class Excel {

    //Open workbook doc
    public XSSFWorkbook getWorkbook(String path) {
        XSSFWorkbook workbook;

        try {
            workbook = new XSSFWorkbook(path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return workbook;
    }

    //Read Excel document
    public Map<Integer, List<String>> getWorkbookSheet(XSSFWorkbook workbook) {

        Sheet sheet = workbook.getSheetAt(0);

        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                    case STRING:
                        data.get(i).add(cell.getRichStringCellValue().getString());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            data.get(i).add(cell.getDateCellValue() + "");
                        } else {
                            data.get(i).add(cell.getNumericCellValue() + "");
                        }
                        break;
                    case BOOLEAN:
                        data.get(i).add(cell.getBooleanCellValue() + "");
                        break;
                    case FORMULA:
                        data.get(i).add(cell.getCellFormula() + "");
                        break;
                    default:
                        data.get(i).add(" ");
                }
            }
            i++;
        }
        return data;
    }


    //Create a separate sheet
    //headers are name and age by default, title of sheet is persons
    public XSSFWorkbook CreateWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Text Analytics");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        headerStyle.setFont(font);

        Cell headerCell = header.createCell(0);
        headerCell.setCellValue("Word");
        headerCell.setCellStyle(headerStyle);

        headerCell = header.createCell(1);
        headerCell.setCellValue("Count");
        headerCell.setCellStyle(headerStyle);

        //Set style
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);

        //set cell values
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("climate");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue(14);
        cell.setCellStyle(style);

        //Write the content to current directory and close the workbook
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(fileLocation);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        try {
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return workbook;
    }


    public static void main(String[] args) throws IOException {
        //Open Excel document and grab first sheet
        String fileLocation = new File("").getAbsolutePath() + "\\TestDoc.xlsx";
        Excel xcel = new Excel();
        XSSFWorkbook wb = xcel.getWorkbook(fileLocation);
        System.out.println("Workbook found");
        //print first sheet
        Map<Integer, List<String>> sheet = xcel.getWorkbookSheet(wb);
        System.out.println("sheet is " + sheet);
        //convert txt file to String array
        ConvertTxtFile();

        //create new workbook
        XSSFWorkbook newWB = xcel.CreateWorkbook();
        System.out.println("New workbook created");
        //print new workbook at first sheet
        Map<Integer, List<String>> newSheet = xcel.getWorkbookSheet(newWB);
        System.out.println("newSheet is " + newSheet);

    }
}



        //TEST
//        public class ExcelIntegrationTest {
//
//            private ExcelPOIHelper excelPOIHelper;
//            private static String FILE_NAME = "temp.xlsx";
//            private String fileLocation;
//
//            @Before
//            public void generateExcelFile() throws IOException {
//                File currDir = new File(".");
//                String path = currDir.getAbsolutePath();
//                fileLocation = path.substring(0, path.length() - 1) + FILE_NAME;
//
//                excelPOIHelper = new ExcelPOIHelper();
//                excelPOIHelper.writeExcel();
//            }
//
//            @Test
//            public void whenParsingPOIExcelFile_thenCorrect() throws IOException {
//                Map<Integer, List<String>> data
//                        = excelPOIHelper.readExcel(fileLocation);
//
//                assertEquals("Name", data.get(0).get(0));
//                assertEquals("Age", data.get(0).get(1));
//
//                assertEquals("John Smith", data.get(1).get(0));
//                assertEquals("20", data.get(1).get(1));
//            }
//        }



