# Excel to CSV Converter

A command-line tool to convert between CSV and Excel (XLSX) files using Java and Apache POI.

## Features

- Convert CSV files to Excel (XLSX) format
- Convert Excel (XLSX) files to CSV format
- Support for UTF-8 and custom character encodings
- Configurable CSV delimiters
- Automatic numeric value detection
- Proper handling of special characters in CSV

## Requirements

- Java 21 or higher
- Maven 3.6 or higher

## Building

```bash
mvn clean package
```

This will create two JAR files in the `target` directory:
- `exceltocsv-1.0.0.jar` - Basic JAR
- `exceltocsv-1.0.0-jar-with-dependencies.jar` - Executable JAR with all dependencies

## Usage

### Basic Usage

Convert CSV to Excel:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar csv2excel input.csv output.xlsx
```

Convert Excel to CSV:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar excel2csv input.xlsx output.csv
```

### Auto-generate Output Filename

If you don't specify an output file, it will be automatically generated:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar csv2excel input.csv
# Creates: input.xlsx

java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar excel2csv data.xlsx
# Creates: data.csv
```

### Custom Options

Specify character encoding:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar csv2excel input.csv output.xlsx --encoding=UTF-8
```

Specify CSV delimiter:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar csv2excel input.csv output.xlsx --delimiter=;
```

Combine options:
```bash
java -jar target/exceltocsv-1.0.0-jar-with-dependencies.jar excel2csv data.xlsx output.csv --encoding=UTF-8 --delimiter=,
```

## Running Tests

```bash
mvn test
```

## License

MIT License
