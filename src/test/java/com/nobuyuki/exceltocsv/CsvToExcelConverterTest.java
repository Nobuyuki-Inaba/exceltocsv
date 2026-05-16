package com.nobuyuki.exceltocsv;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

    @Test
    public void testUpdateExistingSheetByName() throws IOException {
        // Create initial Excel with two sheets
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet s1 = workbook.createSheet("Sheet1");
            s1.createRow(0).createCell(0).setCellValue("original");
            Sheet s2 = workbook.createSheet("Sheet2");
            s2.createRow(0).createCell(0).setCellValue("other-sheet");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        writeToFile(tempCsvFile, "updated,data");
        converter.update(tempCsvFile, tempExcelFile, "Sheet1");

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            assertEquals(2, workbook.getNumberOfSheets());
            // Sheet1 is updated
            Sheet s1 = workbook.getSheet("Sheet1");
            assertNotNull(s1);
            assertEquals("updated", s1.getRow(0).getCell(0).getStringCellValue());
            // Sheet2 is untouched
            Sheet s2 = workbook.getSheet("Sheet2");
            assertNotNull(s2);
            assertEquals("other-sheet", s2.getRow(0).getCell(0).getStringCellValue());
            // Sheet1 is still at index 0
            assertEquals("Sheet1", workbook.getSheetName(0));
        }
    }

    @Test
    public void testUpdateCreatesNewSheetWhenNameNotFound() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1").createRow(0).createCell(0).setCellValue("existing");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        writeToFile(tempCsvFile, "new,sheet,data");
        converter.update(tempCsvFile, tempExcelFile, "NewSheet");

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            assertEquals(2, workbook.getNumberOfSheets());
            assertNotNull(workbook.getSheet("NewSheet"));
            assertEquals("new", workbook.getSheet("NewSheet").getRow(0).getCell(0).getStringCellValue());
            // Original sheet is preserved
            assertEquals("existing", workbook.getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    public void testUpdateBySheetIndex() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Alpha").createRow(0).createCell(0).setCellValue("alpha-data");
            workbook.createSheet("Beta").createRow(0).createCell(0).setCellValue("beta-data");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        writeToFile(tempCsvFile, "replaced");
        converter.update(tempCsvFile, tempExcelFile, 1);

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            assertEquals(2, workbook.getNumberOfSheets());
            // Beta (index 1) replaced
            assertEquals("Beta", workbook.getSheetName(1));
            assertEquals("replaced", workbook.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
            // Alpha (index 0) untouched
            assertEquals("alpha-data", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    public void testUpdatePreservesSheetOrder() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("A").createRow(0).createCell(0).setCellValue("a");
            workbook.createSheet("B").createRow(0).createCell(0).setCellValue("b");
            workbook.createSheet("C").createRow(0).createCell(0).setCellValue("c");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }

        writeToFile(tempCsvFile, "b-updated");
        converter.update(tempCsvFile, tempExcelFile, "B");

        try (Workbook workbook = WorkbookFactory.create(tempExcelFile)) {
            assertEquals("A", workbook.getSheetName(0));
            assertEquals("B", workbook.getSheetName(1));
            assertEquals("C", workbook.getSheetName(2));
            assertEquals("b-updated", workbook.getSheetAt(1).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    public void testUpdateSheetIndexOutOfRange() throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            try (FileOutputStream out = new FileOutputStream(tempExcelFile)) {
                workbook.write(out);
            }
        }
        writeToFile(tempCsvFile, "data");

        try {
            converter.update(tempCsvFile, tempExcelFile, 5);
            fail("Expected IllegalArgumentException");
        } catch (IllegalArgumentException e) {
            assertTrue(e.getMessage().contains("5"));
        }
    }

    private void writeToFile(File file, String content) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(
                new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8))) {
            writer.write(content);
        }
    }
}
