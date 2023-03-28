//package org.excel;
//
//import org.junit.Before;
//import org.junit.Test;
//
//import java.io.File;
//import java.io.FileNotFoundException;
//import java.io.IOException;
//
//import java.util.List;
//
//public class ExcelIntegrationTest {
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
//                excelPOIHelper = new ExcelPOIHelper() {
//                    @Override
//                    public void writeExcel() {
//
//                    }
//
//                    @Override
//                    public void readExcel() {
//
//                    }
//                };
//                excelPOIHelper.writeExcel();
//            }

//            @Test
//            public void whenParsingPOIExcelFile_thenCorrect() throws IOException {
//                Map<Integer, List<String>> data
//                        = excelPOIHelper.readExcel();
//
//                assertEquals("Word", data.get(0).get(0));
//                assertEquals("Count", data.get(0).get(1));
//
//                assertEquals("climate", data.get(1).get(0));
//                assertEquals("14", data.get(1).get(1));
//            }
}
