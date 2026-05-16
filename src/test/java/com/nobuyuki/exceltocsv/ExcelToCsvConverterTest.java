package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.Assert.*;

public class ExcelToCsvConverterTest {

    private ExcelToCsvConverter converter;
    private File tempExcelFile;
    private File tempCsvFile;

    @Before
    public void setUp() throws IOException {
        converter = new ExcelToCsvConverter();
        tempExcelFile = File.createTempFile("test", ".xlsx");
        tempCsvFile = File.createTempFile("test", ".csv");
    }

    @After
    public void tearDown() {
        if (tempExcelFile != null && tempExcelFile.exists()) {
            tempExcelFile.delete();
        }
        if (tempCsvFile != null && tempCsvFile.exists()) {
            tempCsvFile.delete();
        }
    }

    @Test
    public void testBasicExcelToCsvConversion() throws IOException {
        // Create a simple Excel file
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("City");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("John");
            dataRow.createCell(1).setCellValue(30);
            dataRow.createCell(2).setCellValue("New York");

            try (FileOutputStream outputStream = new FileOutputStream(tempExcelFile)) {
                workbook.write(outputStream);
            }
        }

        // Convert to CSV
        converter.convert(tempExcelFile, tempCsvFile);

        // Verify the CSV file
        assertTrue(tempCsvFile.exists());
        assertTrue(tempCsvFile.length() > 0);

        // Read and verify content
        String content = readFile(tempCsvFile);
        String[] lines = content.split("\n");
        assertEquals(2, lines.length);
        assertTrue(lines[0].contains("Name"));
        assertTrue(lines[0].contains("Age"));
        assertTrue(lines[0].contains("City"));
        assertTrue(lines[1].contains("John"));
        assertTrue(lines[1].contains("30"));
        assertTrue(lines[1].contains("New York"));
    }

    @Test
    public void testExcelWithNumericValues() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            Row row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue(123);
            row1.createCell(1).setCellValue(456.78);

            try (FileOutputStream outputStream = new FileOutputStream(tempExcelFile)) {
                workbook.write(outputStream);
            }
        }

        converter.convert(tempExcelFile, tempCsvFile);

        String content = readFile(tempCsvFile);
        assertTrue(content.contains("123"));
        assertTrue(content.contains("456.78"));
    }

    @Test
    public void testCustomDelimiter() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Name");
            row.createCell(1).setCellValue("Age");
            try (FileOutputStream outputStream = new FileOutputStream(tempExcelFile)) {
                workbook.write(outputStream);
            }
        }

        ExcelToCsvConverter semicolonConverter = new ExcelToCsvConverter(StandardCharsets.UTF_8, ";");
        semicolonConverter.convert(tempExcelFile, tempCsvFile);

        String content = readFile(tempCsvFile);
        assertTrue(content.contains("Name;Age"));
        assertFalse(content.contains("Name,Age"));
    }

    @Test
    public void testCustomEncoding() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("名前");
            row.createCell(1).setCellValue("年齢");
            try (FileOutputStream outputStream = new FileOutputStream(tempExcelFile)) {
                workbook.write(outputStream);
            }
        }

        ExcelToCsvConverter utf8Converter = new ExcelToCsvConverter(StandardCharsets.UTF_8, ",");
        utf8Converter.convert(tempExcelFile, tempCsvFile);

        String content = readFile(tempCsvFile);
        assertTrue(content.contains("名前"));
        assertTrue(content.contains("年齢"));
    }

    @Test
    public void testExcelWithSpecialCharacters() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Value with, comma");
            row.createCell(1).setCellValue("Value with \"quotes\"");

            try (FileOutputStream outputStream = new FileOutputStream(tempExcelFile)) {
                workbook.write(outputStream);
            }
        }

        converter.convert(tempExcelFile, tempCsvFile);

        String content = readFile(tempCsvFile);
        // Values with special characters should be properly escaped
        assertTrue(content.contains("\"Value with, comma\""));
        assertTrue(content.contains("\"Value with \"\"quotes\"\"\""));
    }

    @Test
    public void testListSheets() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            workbook.createSheet("売上データ");
            workbook.createSheet("マスタ");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        List<String> names = converter.listSheets(tempExcelFile);

        assertEquals(3, names.size());
        assertEquals("Sheet1", names.get(0));
        assertEquals("売上データ", names.get(1));
        assertEquals("マスタ", names.get(2));
    }

    @Test
    public void testConvertBySheetName() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet s1 = workbook.createSheet("Sheet1");
            s1.createRow(0).createCell(0).setCellValue("from-sheet1");

            Sheet s2 = workbook.createSheet("売上データ");
            s2.createRow(0).createCell(0).setCellValue("from-urriage");

            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        converter.convert(tempExcelFile, tempCsvFile, "売上データ");

        String content = readFile(tempCsvFile);
        assertTrue(content.contains("from-urriage"));
        assertFalse(content.contains("from-sheet1"));
    }

    @Test
    public void testConvertBySheetIndex() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet s0 = workbook.createSheet("Sheet1");
            s0.createRow(0).createCell(0).setCellValue("index0");

            Sheet s1 = workbook.createSheet("Sheet2");
            s1.createRow(0).createCell(0).setCellValue("index1");

            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        converter.convert(tempExcelFile, tempCsvFile, 1);

        String content = readFile(tempCsvFile);
        assertTrue(content.contains("index1"));
        assertFalse(content.contains("index0"));
    }

    @Test
    public void testConvertBySheetNameNotFound() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        try {
            converter.convert(tempExcelFile, tempCsvFile, "NonExistent");
            fail("Expected IllegalArgumentException");
        } catch (IllegalArgumentException e) {
            assertTrue(e.getMessage().contains("NonExistent"));
        }
    }

    @Test
    public void testConvertBySheetIndexOutOfRange() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        try {
            converter.convert(tempExcelFile, tempCsvFile, 5);
            fail("Expected IllegalArgumentException");
        } catch (IllegalArgumentException e) {
            assertTrue(e.getMessage().contains("5"));
        }
    }

    private String readFile(File file) throws IOException {
        StringBuilder content = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(
                new InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8))) {
            String line;
            while ((line = reader.readLine()) != null) {
                content.append(line).append("\n");
            }
        }
        return content.toString();
    }
}
