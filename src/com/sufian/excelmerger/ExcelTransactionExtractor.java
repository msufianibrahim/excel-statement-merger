package com.sufian.excelmerger;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

public class ExcelTransactionExtractor {

    public static List<Transaction> extractTransactionsFromExcel(File file, String selectedMonth, String fileName) throws IOException {
        List<Transaction> transactions = new ArrayList<>();
        SimpleDateFormat[] dateFormats = {
                new SimpleDateFormat("dd/MM/yyyy"),
                new SimpleDateFormat("dd/MM/yy"),
                new SimpleDateFormat("dd MMM"),
                new SimpleDateFormat("dd/MM"),
                new SimpleDateFormat("dd-MM-yyyy"),
                new SimpleDateFormat("dd-MM")
        };
        

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming transactions are in the first sheet

            for (Row row : sheet) {
                Cell dateCell = row.getCell(0);
                Cell descriptionCell = row.getCell(1);
                Cell amountCell = row.getCell(2);

                if (dateCell != null && descriptionCell != null && amountCell != null) {
                    Date date = parseDate(dateCell.getStringCellValue(), dateFormats);
                    
                    if (date != null) {
                        String month = new SimpleDateFormat("MMMM").format(date);
                        String dateStr = Transaction.getFormattedDate(date);
                        if (selectedMonth.equalsIgnoreCase(month)) {
                            String description = descriptionCell.getStringCellValue();
                            Double amount = amountCell.getNumericCellValue();

                            Transaction transaction = new Transaction(dateStr, description, amount, fileName);
                            transactions.add(transaction);
                        }
                    }
                }
            }
        }
        
        return transactions;
    }

    private static Date parseDate(String dateString, SimpleDateFormat[] dateFormats) {
        for (SimpleDateFormat dateFormat : dateFormats) {
            try {
                return dateFormat.parse(dateString);
            } catch (ParseException ignored) {
            }
        }
        return null;
    }

    public static List<Transaction> extractFromAllExcel(String folderPath, String selectedMonth) {
        List<Transaction> allTransactions = new ArrayList<>();
        try {
            Files.walk(Paths.get(folderPath))
                    .filter(Files::isRegularFile)
                    .filter(path -> !path.getFileName().toString().toLowerCase().contains(selectedMonth.toLowerCase()))
                    .filter(path -> path.toString().toLowerCase().endsWith(".xlsx")) // Filter Excel files
                    .forEach(path -> {
                        try {
                            List<Transaction> transactions = extractTransactionsFromExcel(path.toFile(), selectedMonth, path.getFileName().toString());
                            allTransactions.addAll(transactions); // Add transactions to the list
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    });
        } catch (IOException e) {
            e.printStackTrace();
        }
        Collections.sort(allTransactions);
        return allTransactions;
    	
    }
    
    public static class Transaction implements Comparable<Transaction> {
        private String date;
        private String description;
        private double amount;
        private String fileName;

        public Transaction(String date, String description, double amount, String fileName) {
            this.setDate(date);
            this.setDescription(description);
            this.setAmount(amount);
            this.setFileName(fileName);
        }
        
        public static String getFormattedDate(Date date) {
        	SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM");
        	return dateFormat.format(date);
        }

		public String getDate() {
			return date;
		}

		public void setDate(String date) {
			this.date = date;
		}

		public String getDescription() {
			return description;
		}

		public void setDescription(String description) {
			this.description = description;
		}

		public double getAmount() {
			return amount;
		}

		public void setAmount(double amount) {
			this.amount = amount;
		}

		@Override
		public int compareTo(Transaction other) {
			return this.date.compareTo(other.date);
		}

		public String getFileName() {
			return fileName;
		}

		public void setFileName(String fileName) {
			this.fileName = fileName;
		}
    }

}