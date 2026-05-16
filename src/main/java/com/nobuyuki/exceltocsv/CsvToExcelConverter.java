package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

/**
 * Converts CSV files to Excel (XLSX) format.
 */
public class CsvToExcelConverter {

    private Charset encoding;
    private String delimiter;

    public CsvToExcelConverter() {
        this.encoding = StandardCharsets.UTF_8;
        this.delimiter = ",";
    }

    public CsvToExcelConverter(Charset encoding, String delimiter) {
        this.encoding = encoding;
        this.delimiter = delimiter;
    }

    public void convert(File csvFile, File excelFile) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(excelFile)) {
            Sheet sheet = workbook.createSheet("Sheet1");
            writeCsvToSheet(csvFile, sheet);
            workbook.write(outputStream);
        }
    }

    /**
     * Opens an existing Excel file and replaces the named sheet with CSV data.
     * If the sheet does not exist it is created at the end of the workbook.
     */
    public void update(File csvFile, File excelFile, String sheetName) throws IOException {
        byte[] fileBytes = Files.readAllBytes(excelFile.toPath());
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(fileBytes))) {
            int existingIndex = workbook.getSheetIndex(sheetName);
            if (existingIndex >= 0) {
                workbook.removeSheetAt(existingIndex);
                Sheet sheet = workbook.createSheet(sheetName);
                workbook.setSheetOrder(sheetName, existingIndex);
                writeCsvToSheet(csvFile, sheet);
            } else {
                Sheet sheet = workbook.createSheet(sheetName);
                writeCsvToSheet(csvFile, sheet);
            }
            try (FileOutputStream out = new FileOutputStream(excelFile)) {
                workbook.write(out);
            }
        }
    }

    /**
     * Opens an existing Excel file and replaces the sheet at the given zero-based index with CSV data.
     */
    public void update(File csvFile, File excelFile, int sheetIndex) throws IOException {
        byte[] fileBytes = Files.readAllBytes(excelFile.toPath());
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(fileBytes))) {
            if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
                throw new IllegalArgumentException("Sheet index out of range: " + sheetIndex);
            }
            String sheetName = workbook.getSheetName(sheetIndex);
            workbook.removeSheetAt(sheetIndex);
            Sheet sheet = workbook.createSheet(sheetName);
            workbook.setSheetOrder(sheetName, sheetIndex);
            writeCsvToSheet(csvFile, sheet);
            try (FileOutputStream out = new FileOutputStream(excelFile)) {
                workbook.write(out);
            }
        }
    }

    private void writeCsvToSheet(File csvFile, Sheet sheet) throws IOException {
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(csvFile), encoding))) {
            String line;
            int rowNum = 0;
            while ((line = reader.readLine()) != null) {
                Row row = sheet.createRow(rowNum++);
                String[] values = parseCsvLine(line);
                for (int i = 0; i < values.length; i++) {
                    Cell cell = row.createCell(i);
                    String value = values[i];
                    if (isNumeric(value)) {
                        cell.setCellValue(Double.parseDouble(value));
                    } else {
                        cell.setCellValue(value);
                    }
                }
            }
        }
    }

    private String[] parseCsvLine(String line) {
        return line.split(delimiter, -1);
    }

    private boolean isNumeric(String str) {
        if (str == null || str.isEmpty()) {
            return false;
        }
        try {
            Double.parseDouble(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public void setEncoding(Charset encoding) {
        this.encoding = encoding;
    }

    public void setDelimiter(String delimiter) {
        this.delimiter = delimiter;
    }
}
