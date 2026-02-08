package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

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

    /**
     * Converts a CSV file to Excel format.
     *
     * @param csvFile   Input CSV file
     * @param excelFile Output Excel file
     * @throws IOException if an I/O error occurs
     */
    public void convert(File csvFile, File excelFile) throws IOException {
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(csvFile), encoding));
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(excelFile)) {

            Sheet sheet = workbook.createSheet("Sheet1");
            String line;
            int rowNum = 0;

            while ((line = reader.readLine()) != null) {
                Row row = sheet.createRow(rowNum++);
                String[] values = parseCsvLine(line);

                for (int i = 0; i < values.length; i++) {
                    Cell cell = row.createCell(i);
                    String value = values[i];

                    // Try to parse as number
                    if (isNumeric(value)) {
                        cell.setCellValue(Double.parseDouble(value));
                    } else {
                        cell.setCellValue(value);
                    }
                }
            }

            workbook.write(outputStream);
        }
    }

    /**
     * Parses a CSV line handling quoted values.
     */
    private String[] parseCsvLine(String line) {
        return line.split(delimiter, -1);
    }

    /**
     * Checks if a string is numeric.
     */
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
