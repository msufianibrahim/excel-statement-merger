package com.sufian.excelmerger;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.sufian.excelmerger.ExcelTransactionExtractor.Transaction;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelCreator {

    public static void createExcelFile(List<Transaction> transactions, String filePath) throws IOException {
        // Create a new workbook
        try (Workbook workbook = new XSSFWorkbook()) {
            // Create a sheet
            Sheet sheet = workbook.createSheet("Transactions");

            // Create a header row
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Date", "Description", "Amount"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Create data rows
            int rowNum = 1;
            CellStyle decimalStyle = workbook.createCellStyle();
            DataFormat dataFormat = workbook.createDataFormat();
            decimalStyle.setDataFormat(dataFormat.getFormat("#,##0.00")); // Format for two decimal points
            for (Transaction transaction : transactions) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(transaction.getDate());
                row.createCell(1).setCellValue(transaction.getDescription());

                Cell amountCell = row.createCell(2);
                amountCell.setCellValue(transaction.getAmount());
                amountCell.setCellStyle(decimalStyle); // Apply the decimal style to the amount column
            }

            // Auto-size columns
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
        }
    }
}