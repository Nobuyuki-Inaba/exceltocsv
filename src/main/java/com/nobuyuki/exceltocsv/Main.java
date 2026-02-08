package com.nobuyuki.exceltocsv;

import java.io.File;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

/**
 * Command-line interface for CSV to Excel conversion.
 */
public class Main {

    public static void main(String[] args) {
        if (args.length < 2) {
            printUsage();
            System.exit(1);
        }

        String mode = args[0];
        String inputPath = args[1];
        String outputPath = args.length > 2 ? args[2] : null;
        Charset encoding = StandardCharsets.UTF_8;
        String delimiter = ",";

        // Parse optional arguments
        for (int i = 3; i < args.length; i++) {
            if (args[i].startsWith("--encoding=")) {
                String encodingName = args[i].substring(11);
                try {
                    encoding = Charset.forName(encodingName);
                } catch (Exception e) {
                    System.err.println("Invalid encoding: " + encodingName);
                    System.exit(1);
                }
            } else if (args[i].startsWith("--delimiter=")) {
                delimiter = args[i].substring(12);
            }
        }

        File inputFile = new File(inputPath);
        if (!inputFile.exists()) {
            System.err.println("Input file does not exist: " + inputPath);
            System.exit(1);
        }

        try {
            if ("csv2excel".equalsIgnoreCase(mode)) {
                if (outputPath == null) {
                    outputPath = changeFileExtension(inputPath, "xlsx");
                }
                File outputFile = new File(outputPath);
                
                CsvToExcelConverter converter = new CsvToExcelConverter(encoding, delimiter);
                converter.convert(inputFile, outputFile);
                
                System.out.println("Successfully converted CSV to Excel: " + outputFile.getAbsolutePath());
            } else if ("excel2csv".equalsIgnoreCase(mode)) {
                if (outputPath == null) {
                    outputPath = changeFileExtension(inputPath, "csv");
                }
                File outputFile = new File(outputPath);
                
                ExcelToCsvConverter converter = new ExcelToCsvConverter(encoding, delimiter);
                converter.convert(inputFile, outputFile);
                
                System.out.println("Successfully converted Excel to CSV: " + outputFile.getAbsolutePath());
            } else {
                System.err.println("Invalid mode: " + mode);
                printUsage();
                System.exit(1);
            }
        } catch (Exception e) {
            System.err.println("Error during conversion: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.out.println("Usage: java -jar exceltocsv.jar <mode> <input> [output] [options]");
        System.out.println();
        System.out.println("Modes:");
        System.out.println("  csv2excel  - Convert CSV to Excel (XLSX)");
        System.out.println("  excel2csv  - Convert Excel (XLSX) to CSV");
        System.out.println();
        System.out.println("Arguments:");
        System.out.println("  input      - Input file path");
        System.out.println("  output     - Output file path (optional, auto-generated if not specified)");
        System.out.println();
        System.out.println("Options:");
        System.out.println("  --encoding=<charset>   - Character encoding (default: UTF-8)");
        System.out.println("  --delimiter=<char>     - CSV delimiter (default: ,)");
        System.out.println();
        System.out.println("Examples:");
        System.out.println("  java -jar exceltocsv.jar csv2excel input.csv output.xlsx");
        System.out.println("  java -jar exceltocsv.jar excel2csv input.xlsx output.csv --encoding=UTF-8");
        System.out.println("  java -jar exceltocsv.jar csv2excel data.csv --delimiter=;");
    }

    private static String changeFileExtension(String filePath, String newExtension) {
        int lastDot = filePath.lastIndexOf('.');
        if (lastDot > 0) {
            return filePath.substring(0, lastDot + 1) + newExtension;
        }
        return filePath + "." + newExtension;
    }
}
