# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Build JAR
mvn clean package

# Build without running tests
mvn clean package -DskipTests

# Run tests
mvn test

# Run a single test class
mvn test -Dtest=ExcelToCsvConverterTest
```

## Architecture

A library for bidirectional conversion between Excel (XLSX) and CSV formats. Java 21, Maven, Apache POI 5.2.3.

**Two source files**:
- [ExcelToCsvConverter.java](src/main/java/com/nobuyuki/exceltocsv/ExcelToCsvConverter.java) — Reads the first sheet of an XLSX file via Apache POI, converts cells by type (numeric, date, formula, boolean), and writes RFC-compliant CSV with proper quoting and escaping.
- [CsvToExcelConverter.java](src/main/java/com/nobuyuki/exceltocsv/CsvToExcelConverter.java) — Reads CSV line-by-line, auto-detects numeric cells, and writes a single-sheet XLSX workbook. **Note**: does not handle quoted/escaped CSV values when parsing input.

Both converters accept configurable encoding (`Charset`) and delimiter via their constructors. Default encoding: UTF-8. Default delimiter: `,`.

**Tests**: JUnit 4, two test classes mirroring the two converters.
