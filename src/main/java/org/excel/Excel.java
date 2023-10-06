package org.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

import static org.excel.Utils.ConvertTxtFile;

public class Excel {
    //test
    //Open workbook doc in excel
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

//        Sheet sheet = workbook.getSheetAt(0);
        //Get all sheets
        Map<Integer, List<String>> data = new HashMap<>();
        int i = 0;

        Iterator<Sheet> sheetIterator = workbook.iterator();
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
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
        }
        return data;
    }

    //Text analytics
    public Map<String, Integer> CountWordFrequency(String[] arr, Map<Integer, List<String>> data){
        //convert string array to hashmap with count for each word starting at 0
        Map<String, Integer> hashMap = new HashMap<String, Integer>();
        for(int i = 0; i < arr.length; i++) {
            hashMap.put(arr[i], 0);
        }
            //loop through each column
                for (List column : data.values()) {
                   // loop through each text fragment object
                    for (Object text : column) {
                        //convert objects to string array
                        String[] wordArray = text.toString().toLowerCase().split(" ");
                        for (String word : wordArray){
                            //increment count by 1 in hashmap for each match found
                            if(hashMap.containsKey(word)){
                                hashMap.put(word, hashMap.get(word) + 1);
                            }
                        }
                    }
                }

            return hashMap;

    }


    //Create a separate sheet with text analytics data
    public XSSFWorkbook CreateWorkbook(Map<String, Integer> hashMap) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Text Analytics");
        sheet.setColumnWidth(0, 6000);
        sheet.setColumnWidth(1, 4000);

        Row header = sheet.createRow(0);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Lato");
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

        //Set cell values
        int i = 2;
        int j = 0;
        for (Map.Entry<String, Integer> entry: hashMap.entrySet()) {

            Row row = sheet.createRow(i);
            Cell cell = row.createCell(j);
            cell.setCellValue(entry.getKey());
            cell.setCellStyle(style);

            cell = row.createCell(j + 1);
            cell.setCellValue(entry.getValue());
            cell.setCellStyle(style);

            i++;

            }

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
        Path path = Paths.get("words.txt");
        //Open Excel document and grab first sheet
        String fileLocation = new File("").getAbsolutePath() + "\\TestDoc.xlsx";
        Excel xcel = new Excel();
        XSSFWorkbook wb = xcel.getWorkbook(fileLocation);
        System.out.println("Workbook found");
        //print workbook sheets
        Map<Integer, List<String>> sheet = xcel.getWorkbookSheet(wb);
        System.out.println("sheet is " + sheet);
        //convert txt file to String array
        String[] arr = ConvertTxtFile(path);
        //do text analytics
        Map<String, Integer> hashMap = xcel.CountWordFrequency(arr, sheet);
        //create new workbook
        XSSFWorkbook newWB = xcel.CreateWorkbook(hashMap);
        System.out.println("New workbook created");
        //print new workbook
        Map<Integer, List<String>> newSheet = xcel.getWorkbookSheet(newWB);
        System.out.println("newSheet is " + newSheet);

    }
}




