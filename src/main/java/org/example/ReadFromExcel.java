package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class ReadFromExcel {
    public static void main(String[] args) {
        String filePath = "C:\\Practice\\intellij_Workplace\\Exam14Jul2024\\Employee File.xlsx";
        Scanner scanner = new Scanner(System.in);

        System.out.print("1002: ");


        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(file)) {

            Sheet sheet = workbook.getSheetAt(0);

            boolean found = false;
            String searchId = String.valueOf(1002);
            for (Row row : sheet) {

                if (row.getRowNum() == 0) {
                    continue;
                }


                Cell cell = row.getCell(0);
                if (cell != null && cell.getStringCellValue().equals(searchId)) {
                    found = true;
                    System.out.println("Employee Details:");
                    System.out.println("Employee ID: " + row.getCell(0));
                    System.out.println("Employee Name: " + row.getCell(1));
                    System.out.println("Employee Salary: " + row.getCell(2));
                    System.out.println("Employee Mobile: " + row.getCell(3));
                    System.out.println("Employee City: " + row.getCell(4));
                    System.out.println("Manager ID: " + row.getCell(5));
                    System.out.println("Employee Dept: " + row.getCell(6));
                    System.out.println("Employee Share (%): " + row.getCell(7));
                    break;
                }
            }

            if (!found) {
                System.out.println("Employee ID " + searchId + " not found in the Excel file.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
