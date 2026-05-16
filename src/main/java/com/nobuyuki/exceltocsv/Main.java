package com.nobuyuki.exceltocsv;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        if (args.length < 1) {
            printUsage();
            System.exit(1);
        }

        String command = args[0];
        try {
            switch (command) {
                case "list-sheets":
                    handleListSheets(args);
                    break;
                case "excel2csv":
                    handleExcelToCsv(args);
                    break;
                case "csv2excel":
                    handleCsvToExcel(args);
                    break;
                default:
                    System.err.println("Unknown command: " + command);
                    printUsage();
                    System.exit(1);
            }
        } catch (IllegalArgumentException e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        } catch (IOException e) {
            System.err.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }

    private static void handleListSheets(String[] args) throws IOException {
        if (args.length < 2) {
            System.err.println("Usage: list-sheets <input.xlsx>");
            System.exit(1);
        }
        File inputFile = new File(args[1]);
        if (!inputFile.exists()) {
            System.err.println("Error: File not found: " + args[1]);
            System.exit(1);
        }
        List<String> sheets = new ExcelToCsvConverter().listSheets(inputFile);
        for (String name : sheets) {
            System.out.println(name);
        }
    }

    private static void handleExcelToCsv(String[] args) throws IOException {
        if (args.length < 3) {
            System.err.println("Usage: excel2csv <input.xlsx> <output.csv> [--sheet=<name>] [--sheet-index=<n>]");
            System.exit(1);
        }
        File inputFile = new File(args[1]);
        File outputFile = new File(args[2]);

        if (!inputFile.exists()) {
            System.err.println("Error: File not found: " + args[1]);
            System.exit(1);
        }

        String sheetName = null;
        Integer sheetIndex = null;

        for (int i = 3; i < args.length; i++) {
            if (args[i].startsWith("--sheet=")) {
                sheetName = args[i].substring("--sheet=".length());
            } else if (args[i].startsWith("--sheet-index=")) {
                try {
                    sheetIndex = Integer.parseInt(args[i].substring("--sheet-index=".length()));
                } catch (NumberFormatException e) {
                    System.err.println("Error: Invalid sheet index: " + args[i]);
                    System.exit(1);
                }
            }
        }

        ExcelToCsvConverter converter = new ExcelToCsvConverter();
        if (sheetName != null) {
            converter.convert(inputFile, outputFile, sheetName);
        } else if (sheetIndex != null) {
            converter.convert(inputFile, outputFile, sheetIndex);
        } else {
            converter.convert(inputFile, outputFile);
        }
    }

    private static void handleCsvToExcel(String[] args) throws IOException {
        if (args.length < 3) {
            System.err.println("Usage: csv2excel <input.csv> <output.xlsx> [--sheet=<name>] [--sheet-index=<n>] [--update]");
            System.exit(1);
        }
        File inputFile = new File(args[1]);
        File outputFile = new File(args[2]);

        if (!inputFile.exists()) {
            System.err.println("Error: File not found: " + args[1]);
            System.exit(1);
        }

        String sheetName = null;
        Integer sheetIndex = null;
        boolean update = false;

        for (int i = 3; i < args.length; i++) {
            if (args[i].startsWith("--sheet=")) {
                sheetName = args[i].substring("--sheet=".length());
            } else if (args[i].startsWith("--sheet-index=")) {
                try {
                    sheetIndex = Integer.parseInt(args[i].substring("--sheet-index=".length()));
                } catch (NumberFormatException e) {
                    System.err.println("Error: Invalid sheet index: " + args[i]);
                    System.exit(1);
                }
            } else if (args[i].equals("--update")) {
                update = true;
            }
        }

        CsvToExcelConverter converter = new CsvToExcelConverter();
        if (update) {
            if (!outputFile.exists()) {
                System.err.println("Error: --update requires an existing output file: " + args[2]);
                System.exit(1);
            }
            if (sheetName != null) {
                converter.update(inputFile, outputFile, sheetName);
            } else if (sheetIndex != null) {
                converter.update(inputFile, outputFile, sheetIndex);
            } else {
                converter.update(inputFile, outputFile, 0);
            }
        } else {
            converter.convert(inputFile, outputFile);
        }
    }

    private static void printUsage() {
        System.err.println("Usage:");
        System.err.println("  list-sheets <input.xlsx>");
        System.err.println("  excel2csv <input.xlsx> <output.csv> [--sheet=<name>] [--sheet-index=<n>]");
        System.err.println("  csv2excel <input.csv> <output.xlsx> [--sheet=<name>] [--sheet-index=<n>] [--update]");
    }
}
