package excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class WriteExcel {

    /* STEPS
       create a data set which needs to be added to excel
       initialize FileInputStream, create sheet, create headers and set them
       use for loop to iterate the Object[] and set them
       create FileOutputStream and add write worksheet
     */
    @Test
    public void writeDataIntoExcel() {
        Object[][] inputData = {
                {"bipin", "123bipinpassword"},
                {"manju", "123manjupassword"}
                };
        try(FileInputStream inputStream = new FileInputStream("src/test/resources/TestData.xlsx")) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.createSheet("DataStoring");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Username");
            headerRow.createCell(1).setCellValue("Password");

            int rowNumber = 1;
            for(Object[] data : inputData) {
                Row row = sheet.createRow(rowNumber++);
                row.createCell(0).setCellValue((String) data[0]);
                row.createCell(1).setCellValue((String) data[1]);
            }

            try(FileOutputStream outputStream= new FileOutputStream("src/test/resources/TestData.xlsx")){
                workbook.write(outputStream);
            }
        }catch(IOException e) {
            e.printStackTrace();
        }

    }

    @Test
    public void writeDataIntoExcelWithHashmap() {
        Map<String,String> inputData = new HashMap<>();
        inputData.put("bipin","bipin123");
        inputData.put("manju","manju456");

        try(FileInputStream inputStream = new FileInputStream("src/test/resources/TestData.xlsx");
            Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.createSheet("HashMapInputData");
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Uname");
            headerRow.createCell(1).setCellValue("Pwd");

            int rowNumber = 1;
            for(Map.Entry<String, String> dataSet : inputData.entrySet()) {
                Row row = sheet.createRow(rowNumber++);
                row.createCell(0).setCellValue(dataSet.getKey());
                row.createCell(1).setCellValue(dataSet.getValue());
            }

            try(FileOutputStream outputStream = new FileOutputStream("src/test/resources/TestData.xlsx")) {
                workbook.write(outputStream);
            }

        }catch(IOException e) {
            e.printStackTrace();
        }
    }
}
