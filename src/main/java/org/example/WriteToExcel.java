package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

public class WriteToExcel {
    public static void main(String[] args) throws IOException, IOException {
        File file = new File("C:\\Practice\\intellij_Workplace\\Mavin_Project\\src\\main\\resources\\students.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Employee File");

        String empdata[][] =
                {
                        {"Employee_id", "Employee_name", "Emp_Salary", "Emp_mobile", "Emp_city", "Manager_id", "emp_dept","emp_share"},
                        {"1001", "Jack", "1482.45", "0809808008", "NYC", "Null", "Finance", "60"},
                        {"1002", "Joy", "5282.12", "9809808008", "SD", "1001", "Finance", "20"},
                        {"1003", "Nick", "3454.11", "8976876786", "Dayton","1004", "R&D", "30"},
                        {"1004", "Joe", "6482.45", "8809808008", "NYC", "1004", "R&D", "40"},
                        {"1005", "Nick", "5482.45", "5809808008", "CA", "1001", "Finance", "20"},
                        {"1006", "Hyder", "9482.45", "2809808008", "LA", "1005", "Finance", "15"},
                        {"1007", "Harry", "1182.45", "4809808008", "Ohio", "1001", "Finance", "25"},

                };




        for (int rowNum = 0; rowNum < empdata.length; rowNum++) {
            Row row = sheet.createRow(rowNum); // Create a new row

            for (int colNum = 0; colNum < empdata[rowNum].length; colNum++) {
                Cell cell = row.createCell(colNum); // Create a new cell

                cell.setCellValue(empdata[rowNum][colNum]); // Set cell value
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream("Employee File.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Successfully Write back to excel file..");



    }

}

