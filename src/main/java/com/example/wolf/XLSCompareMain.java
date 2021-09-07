package com.example.wolf;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

public class XLSCompareMain {
    public static void main (String[] args) {
        System.out.println("Trying to compare XLSX with hierarchy rows.");

        String dir = System.getProperty("user.dir");
        String sampleName = dir + "\\data\\sample.xlsx";
        String checkName = dir + ".\\data\\check.xlsx";
        String resultName = dir + ".\\data\\result.xlsx";

        if (args.length >= 3) {
            sampleName = args[0];
            checkName = args[1];
            resultName = args[2];
        }
        else {
            System.out.println("Usage: <sample file> <new file to compare> <file with results>");
            System.out.println("You didn't define any parameter, try to default names...");
            System.out.println("Current directory: " + dir);
        }

        doCompare(sampleName, checkName, resultName);
    }

    private static void doCompare(String sampleName, String checkName, String resultName) {

        String fileName = "?";
        try {

            fileName = sampleName;
            System.out.println("Loading sample file " + fileName);
            LinkedHashMap<String, String> sample = readFromExcel(fileName);
            System.out.println("Rows loaded: " + sample.size());

            fileName = checkName;
            System.out.println("Loading check file " + fileName);
            LinkedHashMap<String, String> check = readFromExcel(fileName);
            System.out.println("Rows loaded: " + check.size());

            LinkedHashMap<String, String> deleted = new LinkedHashMap<>();
            LinkedHashMap<String, String> added = new LinkedHashMap<>();

            for (Map.Entry<String, String> item : sample.entrySet()) {
                String key = item.getKey();
                String value = item.getValue();
                if (!check.containsKey(key)) {
                    deleted.put(key, value);
                }
            }

            for (Map.Entry<String, String> item : check.entrySet()) {
                String key = item.getKey();
                String value = item.getValue();
                if (!sample.containsKey(key)) {
                    added.put(key, value);
                }
            }

            System.out.println("Compared.");

            outResult(resultName, added, deleted);

        }
        catch (IOException e) {
            System.out.println("XLSX file read error: " + fileName);
        }

    }

    private static LinkedHashMap<String, String> readFromExcel(String file) throws IOException{

        LinkedHashMap<String, String> array = new LinkedHashMap<>();

        XSSFWorkbook myExcelBook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet myExcelSheet = myExcelBook.getSheetAt(0);

        int rowNum = 2; // Start row (skip header)

        int oldLevel = 0;
        String oldName = "";
        ArrayList<String> path = new ArrayList<>();
        path.add("/");

        do {

            XSSFRow row = myExcelSheet.getRow(rowNum);
            if (row == null) break;

            int level = 0;
            String name = "";

            String currentPath = buildPath(path);

            if (row.getCell(0).getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
                level = (int) row.getCell(0).getNumericCellValue();
            }

            if (row.getCell(1).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                name = row.getCell(1).getStringCellValue();
            }

            if (level == 0) {
                if (name.length() != 0) {
                    System.out.println("ERROR. Row " + rowNum + " has no level but name with value: " + name);
                }
                break;
            }

            if (level != oldLevel) {
                if (level > oldLevel) {
                    if (level - oldLevel > 1) {
                        System.out.println("ERROR. Row " + rowNum + " has level " + level + ". It's not suitable for previous row level " + oldLevel);
                        break;
                    } else {
                        if (oldName.length() > 0) path.add(oldName);
                    }
                }
                else  {
                    for (int i = oldLevel; i > level; i--) {
                        path.remove(i - 1);
                    }
                }
                oldLevel = level;
                currentPath = buildPath(path);
            }

            array.put(currentPath + "|" + name, name);

//            System.out.println("Row: " + rowNum + ": " + level + ": " + name + " | " + currentPath);

            oldName = name;

            rowNum++;

            if (level < 1) break;

        }
        while (true);

        myExcelBook.close();

        return array;

    }

    private static void outResult(String fileName, LinkedHashMap<String, String> added, LinkedHashMap<String, String> deleted) {

        Workbook book = new XSSFWorkbook();

        if (added.size() > 0 || deleted.size() > 0) {

            if (added.size() > 0) {
                Sheet sheet = book.createSheet("Added");
                int rowNum = 0;
                System.out.println("+++ Added rows: " + added.size() + " +++");
                for (Map.Entry<String, String> item : added.entrySet()) {
                    Row row = sheet.createRow(rowNum);
                    Cell path = row.createCell(0);
                    path.setCellValue(getPath(item.getKey()));
                    Cell name = row.createCell(1);
                    name.setCellValue(getName(item.getKey()));
                    rowNum++;
                    System.out.println(path + " > " + name);
                }
            } else {
                System.out.println("There are no added rows.");
            }

            if (deleted.size() > 0) {
                Sheet sheet = book.createSheet("Deleted");
                int rowNum = 0;
                System.out.println("--- Deleted rows: " + deleted.size() + "---");
                for (Map.Entry<String, String> item : deleted.entrySet()) {
                    Row row = sheet.createRow(rowNum);
                    Cell path = row.createCell(0);
                    path.setCellValue(getPath(item.getKey()));
                    Cell name = row.createCell(1);
                    name.setCellValue(getName(item.getKey()));
                    rowNum++;
                    System.out.println(path + " > " + name);
                }
            } else {
                System.out.println("There are no deleted rows.");
            }
            try {
                book.write(new FileOutputStream(fileName));
                book.close();
            }
            catch (IOException e) {
                System.out.println("Error saving file: " + fileName);
            }
        }
    }

    private static String buildPath(ArrayList<String> path) {
        StringBuilder str = new StringBuilder();
        for (int i = 0; i < path.size(); i++) {
            if (i > 1) str.append("/");
            String node = path.get(i);
            if (node.contains("/") && i > 1) str.append("\"");
            str.append(path.get(i));
            if (node.contains("/") && i > 1) str.append("\"");
        }
        return str.toString();
    }

    private static String getName(String item) {
        String name = "";
        if (item != null) {
            int pos = item.lastIndexOf('|');
            if (pos >= 0 && pos < item.length() - 1) name = item.substring(pos + 1);
        }
        return name;
    }

    private static String getPath(String item) {
        String name = "";
        if (item != null) {
            int pos = item.lastIndexOf('|');
            if (pos >= 0) name = item.substring(0, pos);
        }
        return name;
    }
}
