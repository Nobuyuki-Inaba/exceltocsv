package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

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

    /**
     * Converts an Excel file to CSV format.
     *
     * @param excelFile Input Excel file
     * @param csvFile   Output CSV file
     * @throws IOException if an I/O error occurs
     */
    public void convert(File excelFile, File csvFile) throws IOException {
        try (Workbook workbook = WorkbookFactory.create(excelFile);
             BufferedWriter writer = new BufferedWriter(
                     new OutputStreamWriter(new FileOutputStream(csvFile), encoding))) {

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                StringBuilder line = new StringBuilder();
                boolean first = true;

                for (Cell cell : row) {
                    if (!first) {
                        line.append(delimiter);
                    }
                    first = false;

                    String cellValue = getCellValueAsString(cell);
                    line.append(escapeCsvValue(cellValue));
                }

                writer.write(line.toString());
                writer.newLine();
            }
        }
    }

    /**
     * Gets cell value as a string.
     */
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
                    // If it's a whole number, format without decimal places
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

    /**
     * Escapes a CSV value if it contains special characters.
     */
    private String escapeCsvValue(String value) {
        if (value == null) {
            return "";
        }

        // If value contains delimiter, quotes, or newlines, wrap it in quotes
        if (value.contains(delimiter) || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
            // Escape existing quotes by doubling them
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
