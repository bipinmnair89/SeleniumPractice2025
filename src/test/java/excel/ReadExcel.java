package excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcel {

    /* STEPS

       Create a try catch block
       initialize fileInputStream, workbook, sheet, headerRow
       create List of hashmap
       Create for loop for Row, for loop for cell
       create hashmap
       Read data cell by cell in for loop for cell and add it to hashmap and then list
       Extract data from hashmap and print
     */

    @Test
    public void readExcelData() {
        List<Map<String,String>> excelData = new ArrayList<>();
        try(FileInputStream inputStream = new FileInputStream("src/test/resources/TestData.xlsx")) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet sheet = workbook.getSheet("InputData");
            Row headerRow = sheet.getRow(0);
            for(Row row : sheet) {
                if(row.getRowNum() ==0) {
                    continue;
                }
                Map<String,String> eachExcelData = new HashMap<>();
                for(Cell cell : row) {
                    String headerData = headerRow.getCell(cell.getColumnIndex()).getStringCellValue();
                    String valueData = cell.getStringCellValue();
                    eachExcelData.put(headerData,valueData);
                }
                excelData.add(eachExcelData);
            }
            for(Map<String,String> data : excelData) {
                System.out.println("Username - "+data.get("Username")+"\t");
                System.out.println("Password - "+data.get("Password")+"\t");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
