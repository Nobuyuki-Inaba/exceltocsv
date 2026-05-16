package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.*;
import java.nio.charset.StandardCharsets;

import static org.junit.Assert.*;

public class CsvToExcelConverterTest {

    private CsvToExcelConverter converter;
    private File tempCsvFile;
    private File tempExcelFile;

    @Before
    public void setUp() throws IOException {
        converter = new CsvToExcelConverter();
        tempCsvFile = File.createTempFile("test", ".csv");
        tempExcelFile = File.createTempFile("test", ".xlsx");
    }

    @After
    public void tearDown() {
        if (tempCsvFile != null && tempCsvFile.exists()) {
            tempCsvFile.delete();
        }
        if (tempExcelFile != null && tempExcelFile.exists()) {
            tempExcelFile.delete();
        }
    }

    @Test
    public void testBasicCsvToExcelConversion() throws IOException {
        // Create a simple CSV file
        String csvContent = "Name,Age,City\nJohn,30,New York\nJane,25,Los Angeles";
        writeToFile(tempCsvFile, csvContent);

        // Convert to Excel
        converter.convert(tempCsvFile, tempExcelFile);

        // Verify the Excel file
        assertTrue(tempExcelFile.exists());
        assertTrue(tempExcelFile.length() > 0);

        // Verify content
        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(3, sheet.getPhysicalNumberOfRows());

            Row headerRow = sheet.getRow(0);
            assertEquals("Name", headerRow.getCell(0).getStringCellValue());
            assertEquals("Age", headerRow.getCell(1).getStringCellValue());
            assertEquals("City", headerRow.getCell(2).getStringCellValue());

            Row dataRow = sheet.getRow(1);
            assertEquals("John", dataRow.getCell(0).getStringCellValue());
            assertEquals(30.0, dataRow.getCell(1).getNumericCellValue(), 0.001);
            assertEquals("New York", dataRow.getCell(2).getStringCellValue());
        }
    }

    @Test
    public void testEmptyCsvConversion() throws IOException {
        writeToFile(tempCsvFile, "");
        converter.convert(tempCsvFile, tempExcelFile);
        
        assertTrue(tempExcelFile.exists());
    }

    @Test
    public void testCustomDelimiter() throws IOException {
        String csvContent = "Name;Age;City\nJohn;30;Tokyo";
        writeToFile(tempCsvFile, csvContent);

        CsvToExcelConverter semicolonConverter = new CsvToExcelConverter(StandardCharsets.UTF_8, ";");
        semicolonConverter.convert(tempCsvFile, tempExcelFile);

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            assertEquals("Name", headerRow.getCell(0).getStringCellValue());
            assertEquals("Age", headerRow.getCell(1).getStringCellValue());
            assertEquals("City", headerRow.getCell(2).getStringCellValue());
            Row dataRow = sheet.getRow(1);
            assertEquals("John", dataRow.getCell(0).getStringCellValue());
            assertEquals(30.0, dataRow.getCell(1).getNumericCellValue(), 0.001);
            assertEquals("Tokyo", dataRow.getCell(2).getStringCellValue());
        }
    }

    @Test
    public void testCustomEncoding() throws IOException {
        String csvContent = "名前,年齢\n太郎,25";
        writeToFile(tempCsvFile, csvContent);

        CsvToExcelConverter utf8Converter = new CsvToExcelConverter(StandardCharsets.UTF_8, ",");
        utf8Converter.convert(tempCsvFile, tempExcelFile);

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            assertEquals("名前", headerRow.getCell(0).getStringCellValue());
            assertEquals("年齢", headerRow.getCell(1).getStringCellValue());
            Row dataRow = sheet.getRow(1);
            assertEquals("太郎", dataRow.getCell(0).getStringCellValue());
            assertEquals(25.0, dataRow.getCell(1).getNumericCellValue(), 0.001);
        }
    }

    @Test
    public void testNumericValueConversion() throws IOException {
        String csvContent = "Value\n123\n456.78\n-999";
        writeToFile(tempCsvFile, csvContent);

        converter.convert(tempCsvFile, tempExcelFile);

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            assertEquals(123.0, sheet.getRow(1).getCell(0).getNumericCellValue(), 0.001);
            assertEquals(456.78, sheet.getRow(2).getCell(0).getNumericCellValue(), 0.001);
            assertEquals(-999.0, sheet.getRow(3).getCell(0).getNumericCellValue(), 0.001);
        }
    }

    private void writeToFile(File file, String content) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8))) {
            writer.write(content);
        }
    }
}
