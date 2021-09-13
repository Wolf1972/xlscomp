package com.example.wolf;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
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

    /**
     * Compare procedure
     * @param sampleName - source file name
     * @param checkName - file name for compare
     * @param resultName - file name for result
     */
    private static void doCompare(String sampleName, String checkName, String resultName) {

        String fileName = "?";
        try {

            fileName = sampleName;
            System.out.println("Loading sample file " + fileName);
            LinkedHashMap<String, Requirement> sample = readFromExcel(fileName);
            System.out.println("Rows loaded: " + sample.size());

            fileName = checkName;
            System.out.println("Loading check file " + fileName);
            LinkedHashMap<String, Requirement> check = readFromExcel(fileName);
            System.out.println("Rows loaded: " + check.size());

            LinkedHashMap<String, Requirement> deleted = new LinkedHashMap<>();
            LinkedHashMap<String, Requirement> added = new LinkedHashMap<>();

            for (Map.Entry<String, Requirement> item : sample.entrySet()) {
                String key = item.getKey();
                Requirement req = item.getValue();
                if (!check.containsKey(key)) {
                    deleted.put(key, req);
                }
            }

            for (Map.Entry<String, Requirement> item : check.entrySet()) {
                String key = item.getKey();
                Requirement req = item.getValue();
                if (!sample.containsKey(key)) {
                    added.put(key, req);
                }
            }

            System.out.println("Compared.");

            outResult(sampleName, checkName, resultName, sample, check, added, deleted);

        }
        catch (IOException e) {
            System.out.println("XLSX file read error: " + fileName);
        }

    }

    /**
     * Reads one Excel file (first sheet)
     * @param file - file name
     * @return - array with sheet data
     * @throws IOException
     */
    private static LinkedHashMap<String, Requirement> readFromExcel(String file) throws IOException{

        LinkedHashMap<String, Requirement> array = new LinkedHashMap<>();

        XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = book.getSheetAt(0);

        int rowNum = 2; // Start row (skip header)

        int oldLevel = 0;
        String oldName = "";
        ArrayList<String> path = new ArrayList<>();
        path.add("\\");

        do {

            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) break;

            int level = 0;
            String name = "";

            String currentPath = printPath(path);

            if (row.getCell(0).getCellType() == CellType.NUMERIC) {
                level = (int) row.getCell(0).getNumericCellValue();
            }

            if (row.getCell(1).getCellType() == CellType.STRING) {
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
                currentPath = printPath(path);
            }

            Requirement req = new Requirement();
            req.id = currentPath + "|" + name;
            req.loadFromRow(row);
            array.put(req.id, req);

//            System.out.println("Row: " + rowNum + ": " + level + ": " + name + " | " + currentPath);

            oldName = name;

            rowNum++;

            if (level < 1) break;

        }
        while (true);

        book.close();

        return array;

    }

    /**
     * Outputs compare results
     * @param sampleFileName - sample file name
     * @param checkFileName - file name for check
     * @param resultFileName - file name for results
     * @param sample - array with source data
     * @param check - array with data for compare
     * @param added - array with added rows
     * @param deleted - array with deleted rows
     */
    private static void outResult(String sampleFileName, String checkFileName, String resultFileName,
                                  LinkedHashMap<String, Requirement> sample, LinkedHashMap<String, Requirement> check,
                                  LinkedHashMap<String, Requirement> added, LinkedHashMap<String, Requirement> deleted) {

        XSSFWorkbook book = new XSSFWorkbook();

        if (added.size() > 0 || deleted.size() > 0) {

            if (added.size() > 0) {
                System.out.println("+++ Added rows: " + added.size() + " +++");
                XSSFSheet sampleSheet = book.createSheet("Old");
                copySheet(sampleFileName, sampleSheet);
                XSSFSheet addSheet = book.createSheet("Added");
                outOneDiffSheet(addSheet, added);
            } else {
                System.out.println("There are no added rows.");
            }

            if (deleted.size() > 0) {
                XSSFSheet checkSheet = book.createSheet("New");
                copySheet(checkFileName, checkSheet);
                XSSFSheet sheet = book.createSheet("Deleted");
                System.out.println("--- Deleted rows: " + deleted.size() + "---");
                outOneDiffSheet(sheet, deleted);
            } else {
                System.out.println("There are no deleted rows.");
            }
            try {
                book.write(new FileOutputStream(resultFileName));
                book.close();
            }
            catch (IOException e) {
                System.out.println("Error saving file: " + resultFileName);
            }
        }
    }

    /**
     * Outputs one source sheet (sample or check)
     * @param sheet - sheet
     * @param array - array to fill sheet
     */
    private static void outOneSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array) {
        int rowNum = 0;
        for (Map.Entry<String, Requirement> item : array.entrySet()) {
            XSSFRow row = sheet.createRow(rowNum);
            item.getValue().saveToRow(row);
            rowNum++;
        }
    }

    /**
     * Outputs one different sheet (added or deleted rows)
     * @param sheet - sheet
     * @param array - array to fill sheet
     */
    private static void outOneDiffSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array) {
        XSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        XSSFFont font = sheet.getWorkbook().createFont();
        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFont(font);

        int rowNum = 0;
        int level = 0;
        String oldPathStr = "";
        String[] oldPathArr = {};
        for (Map.Entry<String, Requirement> item : array.entrySet()) {
            String pathStr = getPath(item.getKey());
            if (!pathStr.equals(oldPathStr)) {
                String[] pathArr = pathStr.split("\\\\");
                level = 0;
                for (int i = 1; i < pathArr.length; i++) {
/*
                    if (i < oldPathArr.length) { // Do not output common parts of path
                        if (oldPathArr[i].equals(pathArr[i])) continue;
                    }
*/
                    XSSFRow pathRow = sheet.createRow(rowNum);
                    XSSFCell pathCount = pathRow.createCell(0);
                    pathCount.setCellType(CellType.NUMERIC);
                    pathCount.setCellStyle(style);
                    pathCount.setCellValue(i);
                    XSSFCell pathHead = pathRow.createCell(1);
                    pathHead.setCellStyle(style);
                    pathHead.setCellValue(pathArr[i]);
                    rowNum++;
                    level = i;
                }
                oldPathStr = pathStr;
                oldPathArr = pathArr;
            }
            XSSFRow row = sheet.createRow(rowNum);
            XSSFCell path = row.createCell(0);
            path.setCellValue(level + 1);
            XSSFCell name = row.createCell(1);
            name.setCellValue(getName(item.getKey()));
            rowNum++;
            // System.out.println(path + " > " + name);
        }
    }

    /**
     * Builds path string from array with hierarchical rows
     * @param path - array with path
     * @return - string with path divided by "\"
     */
    private static String printPath(ArrayList<String> path) {
        StringBuilder str = new StringBuilder();
        for (int i = 0; i < path.size(); i++) {
            if (i > 1) str.append("\\");
            String node = path.get(i);
            if (node.contains("\\") && i > 1) str.append("\"");
            str.append(path.get(i));
            if (node.contains("\\") && i > 1) str.append("\"");
        }
        return str.toString();
    }

    /**
     * Returns name from full path (path with name), divided by "|"
     * @param item - path with name
     * @return - only name
     */
    private static String getName(String item) {
        String name = "";
        if (item != null) {
            int pos = item.lastIndexOf('|');
            if (pos >= 0 && pos < item.length() - 1) name = item.substring(pos + 1);
        }
        return name;
    }

    /**
     * Returns path from full path (path with name), divided by "|"
     * @param item - path with name
     * @return - only path
     */
    private static String getPath(String item) {
        String name = "";
        if (item != null) {
            int pos = item.lastIndexOf('|');
            if (pos >= 0) name = item.substring(0, pos);
        }
        return name;
    }

    private static void copySheet(String sourceFile, XSSFSheet targetSheet) {

        final int HEADER_LAST_ROW = 1;

        class Group { // Item for rows grouping
            int start;
            int end;
            int level;
            Group(int start, int end, int level) {
                this.start = start;
                this.end = end;
                this.level = level;
            }
        }

        ArrayList<Group> groups = new ArrayList<>();
        groups.add(new Group(3, 23, 2));
        groups.add(new Group(4, 8, 3));
        groups.add(new Group(13, 15, 3));
        groups.add(new Group(17, 18, 3));

        // Common styles for all group levels (outlines)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        try {

            XSSFWorkbook sourceBook = new XSSFWorkbook(new FileInputStream(sourceFile));
            XSSFSheet sourceSheet = sourceBook.getSheetAt(0);
            int oldOutlineLevel = 0;
            int lastRow = sourceSheet.getLastRowNum();
            for (int i = 0; i < lastRow; i++) {
                int outlineLevel = sourceSheet.getRow(i).getOutlineLevel();
                System.out.println("Copying row " + i + " Outline: " + outlineLevel);
                if (outlineLevel > 0) {
                    int specifiedOutlineLevel = (int) sourceSheet.getRow(i).getCell(0).getNumericCellValue();
                    if (outlineLevel + 1 != specifiedOutlineLevel) {
                        System.out.println("Error. Outline level " + outlineLevel + " doesn't suite with level specified in first column: " + specifiedOutlineLevel);
                    }
                }

                ArrayList<XSSFCellStyle> styles = new ArrayList<>();
                if (i <= HEADER_LAST_ROW || !groupStyles.containsKey(outlineLevel)) {
                    // Styles for header or new outline level
                    // Copy style from old cell and apply to new cell: all styles after specified row are common - takes it from array
                    for (int j = 0; j < sourceSheet.getRow(i).getLastCellNum(); j++) {
                        XSSFCell cell = sourceSheet.getRow(i).getCell(j);
                        XSSFCellStyle newCellStyle = targetSheet.getWorkbook().createCellStyle();
                        newCellStyle.cloneStyleFrom(cell.getCellStyle());
                        styles.add(newCellStyle);
                    }
                    if (i > HEADER_LAST_ROW) {
                        groupStyles.put(outlineLevel, styles);
                    }
                }
                else {
                    styles = groupStyles.get(outlineLevel);
                }
                copyRow(sourceSheet, targetSheet, i, i, styles);
            }

            sourceBook.close();

        }
        catch (IOException e) {
            System.out.println("Error while reading source sheet: " + sourceFile);
        }

        targetSheet.setRowSumsBelow(false);
        for (Group group : groups) {
            targetSheet.groupRow(group.start, group.end - 1);
        }

    }

    private static void copyRow(XSSFSheet sourceWorksheet, XSSFSheet targetWorksheet,
                                int sourceRowNum, int targetRowNum,
                                ArrayList<XSSFCellStyle> columnStyles) {
        // Get the source / new row
        XSSFRow newRow = targetWorksheet.getRow(targetRowNum);
        XSSFRow sourceRow = sourceWorksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            targetWorksheet.shiftRows(targetRowNum, targetWorksheet.getLastRowNum(), 1);
        } else {
            newRow = targetWorksheet.createRow(targetRowNum);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            if (i < columnStyles.size()) {
                XSSFCellStyle newCellStyle = columnStyles.get(i);
                newCell.setCellStyle(newCellStyle);
            }

            targetWorksheet.setColumnWidth(i, sourceWorksheet.getColumnWidth(i));

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data value
            CellType type = oldCell.getCellType();
            if (type == CellType.BLANK) {
                newCell.setCellType(type);
                newCell.setCellValue(oldCell.getStringCellValue());
            }
            else if (type == CellType.BOOLEAN) {
                newCell.setCellType(type);
                newCell.setCellValue(oldCell.getBooleanCellValue());
            }
            else if (type == CellType.ERROR) {
                newCell.setCellType(type);
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
            }
            else if (type == CellType.FORMULA) {
                newCell.setCellFormula(oldCell.getCellFormula());
            }
            else if (type == CellType.NUMERIC) {
                newCell.setCellType(type);
                newCell.setCellValue(oldCell.getNumericCellValue());
            }
            else if (type == CellType.STRING) {
                newCell.setCellType(type);
                newCell.setCellValue(oldCell.getRichStringCellValue());
            }
        }

        // If there are any merged regions in the source row, copy to new row
        for (int i = 0; i < sourceWorksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = sourceWorksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                targetWorksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }
}
