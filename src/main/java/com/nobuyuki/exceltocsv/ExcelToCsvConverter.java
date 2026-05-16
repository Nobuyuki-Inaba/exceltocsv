package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * Converts Excel (XLSX) files to CSV format.
 */
public class ExcelToCsvConverter {

    private Charset encoding;
    private String delimiter;

    public ExcelToCsvConverter() {
        this.encoding = StandardCharsets.UTF_8;
        this.delimiter = ",";
    }

    public ExcelToCsvConverter(Charset encoding, String delimiter) {
        this.encoding = encoding;
        this.delimiter = delimiter;
    }

    public List<String> listSheets(File excelFile) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            List<String> names = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                names.add(workbook.getSheetName(i));
            }
            return names;
        }
    }

    public void convert(File excelFile, File csvFile) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            writeSheetToCsv(workbook.getSheetAt(0), csvFile);
        }
    }

    public void convert(File excelFile, File csvFile, int sheetIndex) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
                throw new IllegalArgumentException("Sheet index out of range: " + sheetIndex);
            }
            writeSheetToCsv(workbook.getSheetAt(sheetIndex), csvFile);
        }
    }

    public void convert(File excelFile, File csvFile, String sheetName) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile)) {
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }
            writeSheetToCsv(sheet, csvFile);
        }
    }

    private void writeSheetToCsv(Sheet sheet, File csvFile) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(new FileOutputStream(csvFile), encoding))) {
            for (Row row : sheet) {
                StringBuilder line = new StringBuilder();
                boolean first = true;
                for (Cell cell : row) {
                    if (!first) line.append(delimiter);
                    first = false;
                    line.append(escapeCsvValue(getCellValueAsString(cell)));
                }
                writer.write(line.toString());
                writer.newLine();
            }
        }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double numValue = cell.getNumericCellValue();
                    if (numValue == (long) numValue) {
                        return String.format("%d", (long) numValue);
                    } else {
                        return String.valueOf(numValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf(cell.getNumericCellValue());
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }

    private String escapeCsvValue(String value) {
        if (value == null) {
            return "";
        }
        if (value.contains(delimiter) || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
            value = value.replace("\"", "\"\"");
            return "\"" + value + "\"";
        }
        return value;
    }

    public void setEncoding(Charset encoding) {
        this.encoding = encoding;
    }

    public void setDelimiter(String delimiter) {
        this.delimiter = delimiter;
    }
}
