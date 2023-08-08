package com.excel.file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriter {

    public static void main(String[] args) {
        String filePath = "data.xlsx";
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            String[] headers = {"Name", "Age", "Email"};
            String[][] data = {
                    {"Ajay", "30", "ajay@test.com"},
                    {"Vijay", "28", "vijay@test.com"},
                    {"Sanjay", "35", "sanjay@example.com"},
                    {"Swapnil", "37", "swapnil@example.com"}
            };

            // Write column headers
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Write data rows
            for (int i = 0; i < data.length; i++) {
                Row dataRow = sheet.createRow(i + 1);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = dataRow.createCell(j);
                    cell.setCellValue(data[i][j]);
                }
            }

            // Write the workbook to a file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Data has been written to the Excel file successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
