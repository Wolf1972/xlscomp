package com.example.wolf;

import org.apache.commons.cli.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.*;

public class XLSCompareMain {

    public static void main (String[] args) {

        final Properties properties = new Properties();
        try {
            properties.load(XLSCompareMain.class.getClassLoader().getResourceAsStream("project.properties"));
        }
        catch (IOException e) {
            System.out.println("Error reading resource file.");
        }

        System.out.println("Compare for hierarchical requirements in XLSX. Version " + properties.getProperty("version"));

        String dir = System.getProperty("user.dir"); if (!dir.endsWith("\\")) dir += "\\"; // Current directory

        String oldFile = "data\\old.xlsx";
        String newFile = "data\\new.xlsx";
        String resultFile = "data\\result.xlsx";
        String mxWebFile = "data\\mxweb.xlsx";
        String cmdColumns = "25";
        String mergeMode = "m";

        Options options = new Options();
        options.addOption("o", "old", true, "Old XLSX file for compare");
        options.addOption("n", "new", true, "New XLSX file for compare");
        options.addOption("r", "result", true, "XLSX file for result");
        options.addOption("m", "merge", true, "Merge cells for common rows from old file to new (a - all. e - if only new cell is empty)");
        options.addOption("w", "mxweb", true, "[XLSX file with MxWeb requirements to match]");
        options.addOption("c", "columns", true, "Result columns count (25 columns by default)");
        options.addOption("d", "directory", true, "Common directory for all input and output files");
        options.addOption("h", "help", true, "Print this message");

        CommandLineParser parser = new DefaultParser();
        try {
            CommandLine command = parser.parse(options, args);

            if (command.hasOption('d')) { dir = command.getOptionValue('d'); if (!dir.endsWith("\\")) dir += "\\"; }
            if (command.hasOption('i')) newFile = command.getOptionValue('i');
            if (command.hasOption('i')) resultFile = command.getOptionValue('r');
            if (command.hasOption('o')) oldFile = command.getOptionValue('o');
            if (command.hasOption('m')) mergeMode = command.getOptionValue('m');
            if (command.hasOption('w')) mxWebFile = command.getOptionValue('w');
            if (command.hasOption('c')) cmdColumns = command.getOptionValue('c');
            if (command.hasOption('h')) {
                HelpFormatter formatter = new HelpFormatter();
                formatter.printHelp(properties.getProperty("artifactId"), options);
                return;
            }
        }
        catch (ParseException e) {
            System.out.println("Command line parse exception.");
            HelpFormatter help = new HelpFormatter();
            help.printHelp(XLSCompareMain.class.getSimpleName(), options);
            return;
        }

        String oldName = dir + oldFile;
        String newName = dir + newFile;
        String resultName = dir + resultFile;
        String mxWebName = mxWebFile != null? dir + mxWebFile : "";

        System.out.println("Old requirements: " + oldName);
        System.out.println("New requirements: " + newName);
        System.out.println("Result requirements: " + resultName);
        if (!mxWebName.isEmpty()) System.out.println("MxWeb requirements: " + mxWebName);
        int maxColumn = Integer.parseInt(cmdColumns) - 1; // last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
        if (maxColumn < 0) maxColumn = 13; // set to 13 for public results

        System.out.println();
        doCompare(oldName, newName, mxWebName, resultName, maxColumn, mergeMode);
        System.out.println("Done.");
    }

    /**
     * Compare requirements procedure
     * @param oldName - old file name
     * @param newName - new file name for compare
     * @param mxWebName - file name with mxweb requirements
     * @param resultName - file name for result
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     * @param mergeMode - merge mode ("" - no merge, a - all cells from old file. e - merge only if new cell is empty)
     */
    private static void doCompare(String oldName, String newName, String mxWebName, String resultName, int maxColumn, String mergeMode) {

        String fileName = "?";
        try {

            System.out.println("Trying to compare...");

            fileName = oldName;
            System.out.println("Loading old file " + fileName);
            LinkedHashMap<String, Requirement> oldMap = Requirement.readFromExcel(fileName);
            System.out.println("Rows loaded: " + oldMap.size());

            fileName = newName;
            System.out.println("Loading new file " + fileName);
            LinkedHashMap<String, Requirement> newMap = Requirement.readFromExcel(fileName);
            System.out.println("Rows loaded: " + newMap.size());

            LinkedHashMap<String, Requirement> deletedMap = new LinkedHashMap<>();
            LinkedHashMap<String, Requirement> addedMap = new LinkedHashMap<>();

            LinkedHashMap<String, Requirement> changedMap = new LinkedHashMap<>(); // Map with only changed rows
            LinkedHashMap<String, List<Integer>> changedMapDetails = new LinkedHashMap<>(); // Map with columns has changed for every row has changed

            for (Map.Entry<String, Requirement> item : oldMap.entrySet()) {
                String key = item.getKey();
                Requirement req = item.getValue();
                if (!newMap.containsKey(key)) {
                    deletedMap.put(key, req);
                }
                else {
                    Requirement newReq = newMap.get(key);
                    if (newReq != null) {
                        List<Integer> changedColumns = newReq.compare(req);
                        if (changedColumns != null && changedColumns.size() > 0) {
                            changedMap.put(key, req);
                            changedMapDetails.put(key, changedColumns);
                        }
                    }
                }
            }

            for (Map.Entry<String, Requirement> item : newMap.entrySet()) {
                String key = item.getKey();
                Requirement req = item.getValue();
                if (!oldMap.containsKey(key)) {
                    addedMap.put(key, req);
                }
            }

            System.out.println("Compared.");

            LinkedHashMap<String, Requirement> matchedMap = new LinkedHashMap<>(); // Rows from mxWeb were matched without errors
            LinkedHashMap<Integer, MxRequirement> missedMap = new LinkedHashMap<>(); // Rows from MxWeb requirements were missed when matching

            if (mxWebName != null) {

                System.out.println("Trying to check with MxWeb requirements...");

                fileName = mxWebName;
                System.out.println("Loading mxWeb file " + fileName);
                LinkedHashMap<Integer, MxRequirement> mxwebMap = MxRequirement.readFromExcel(fileName);
                System.out.println("Rows loaded: " + mxwebMap.size());

                for (Map.Entry<String, Requirement> item : newMap.entrySet()) {
                    String mxwebReqId = item.getValue().getReference();
                    if (mxwebReqId != null && mxwebReqId.length() > 0) {
                        boolean isFound = false;
                        if (mxwebReqId.contains("\n")) mxwebReqId = mxwebReqId.replace("\n", " ");
                        List<String> aMReq = new ArrayList<>();
                        if (mxwebReqId.contains(" ")) {
                            String[] array = mxwebReqId.split(" ");
                            aMReq = Arrays.asList(array);
                        } else {
                            aMReq.clear();
                            aMReq.add(mxwebReqId);
                        }
                        for (int k = 0; k < aMReq.size(); k++) {
                            for (Map.Entry<Integer, MxRequirement> mxitem : mxwebMap.entrySet()) {
                                String mxId = mxitem.getValue().getMxwebid();
                                if (mxId.equals(aMReq.get(k))) {
                                    matchedMap.put(item.getKey(), item.getValue());
                                    isFound = true;
                                    // Do not brake after first row found - may be several MxWeb requirements in one cell, we have to mark its all
                                }
                            }
                            if (!isFound) {
                                System.out.println("ERROR. Requirement row " + (item.getValue().getRow() + 1) + " has MxWeb requirement " + aMReq.get(k) + " but has not found in the MxWeb sheet.");
                            }
                        }
                    }
                }

                // Reverse check - try to find MxWeb requirement has missed in our sheet
                for (Map.Entry<Integer, MxRequirement> mxitem : mxwebMap.entrySet()) {
                    String searchFor = mxitem.getValue().getMxwebid();
                    String release = mxitem.getValue().getRelease();
                    if (searchFor != null && !searchFor.isEmpty()) {
                        if (release == null || release.isEmpty()) {
                            System.out.println("WARN. MxWeb requirement " + searchFor + " without release specified.");
                        }
                        boolean isFound = false;
                        for (Map.Entry<String, Requirement> item : newMap.entrySet()) {
                            String mxwebReqId = item.getValue().getReference();
                            if (mxwebReqId != null && !mxwebReqId.isEmpty()) {
                                if (mxwebReqId.contains("\n")) mxwebReqId = mxwebReqId.replace("\n", " ");
                                List<String> aMReq;
                                if (mxwebReqId.contains(" ")) {
                                    String[] array = mxwebReqId.split(" ");
                                    aMReq = Arrays.asList(array);
                                } else {
                                    aMReq = Arrays.asList(mxwebReqId);
                                }
                                for (int k = 0; k < aMReq.size(); k++) {
                                    if (aMReq.get(k).equals(searchFor)) {
                                        matchedMap.put(item.getKey(), item.getValue());
                                        String oldRelease = item.getValue().getSource_req();
                                        if (oldRelease != null && !oldRelease.isEmpty() && !release.isEmpty() && !oldRelease.equals(release)) {
                                            System.out.println("ERROR. Row " + (item.getValue().getRow() + 1) + " requires release " + release + " but already has release " + oldRelease);
                                        }
                                        else {
                                            item.getValue().setSource_req(release);
                                        }
                                        isFound = true;
                                        break;
                                    }
                                }
                                // Do not brake after first row found - may be several requirement rows with the same MxWeb requirement, we have to mark its all
                            }
                        }
                        if (!isFound) {
                            System.out.println("ERROR. MxWeb requirement " + (mxitem.getValue().getMxwebid()) + " has not found in our sheet.");
                            missedMap.put(mxitem.getKey(), mxitem.getValue());
                        }
                    }
                }

                System.out.println("MxWeb requirements checked.");
            }

            outResult(oldName, newName, resultName, maxColumn,
                      addedMap, deletedMap,
                      changedMap, changedMapDetails,
                      matchedMap, missedMap);

        }
        catch (IOException e) {
            System.out.println("XLSX file read error: " + fileName);
        }

    }

    /**
     * Outputs compare results
     * @param oldFileName - old file name
     * @param newFileName - new file name for compare
     * @param resultFileName - file name for results
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     * @param addedMap - map with added rows
     * @param deletedMap - map with deleted rows
     * @param changedMap - map with changed rows
     * @param changedDetailsMap - map with changed details
     * @param matchedMap - map with matched with MaxWeb requirement rows
     * @param missedMap - map with MxWeb requirements rows were missed when matching
     */
    private static void outResult(String oldFileName, String newFileName, String resultFileName, int maxColumn,
                                  LinkedHashMap<String, Requirement> addedMap, LinkedHashMap<String, Requirement> deletedMap,
                                  LinkedHashMap<String, Requirement> changedMap, LinkedHashMap<String, List<Integer>> changedDetailsMap,
                                  LinkedHashMap<String, Requirement> matchedMap, LinkedHashMap<Integer, MxRequirement> missedMap) {

        System.out.println("Output results.");

        XSSFWorkbook book = new XSSFWorkbook();

        if (addedMap.size() > 0 || deletedMap.size() > 0) {

            XSSFSheet currentSheet = book.createSheet("Current");
            XLSUtil.copySheet(newFileName, currentSheet, maxColumn);

            XSSFSheet oldSheet = book.createSheet("Old");
            XLSUtil.copySheet(oldFileName, oldSheet, maxColumn);
//            if (deletedMap.size() > 0) markOneSheet(oldSheet, deletedMap, null, true);
            if (changedMap.size() > 0) markOneSheet(oldSheet, changedMap, changedDetailsMap, null);

            XSSFSheet newSheet = book.createSheet("New");
            XLSUtil.copySheet(newFileName, newSheet, maxColumn);
//            if (addedMap.size() > 0) markOneSheet(newSheet, addedMap, null, false);
            if (changedMap.size() > 0) markOneSheet(newSheet, changedMap, changedDetailsMap, null);

            if (deletedMap.size() > 0) {
                XSSFSheet delSheet = book.createSheet("Deleted");
                copySheetWithFilter(oldSheet, delSheet, deletedMap, true, maxColumn);
                System.out.println("- Deleted rows: " + deletedMap.size());
//                XSSFSheet delSheet2 = book.createSheet("Deleted2");
//                outOneDiffSheet(delSheet2, deletedMap);
            }
            else {
                System.out.println("There are no deleted rows.");
            }

            if (addedMap.size() > 0) {
                XSSFSheet addSheet = book.createSheet("Added");
                copySheetWithFilter(newSheet, addSheet, addedMap, false, maxColumn);
                System.out.println("+ Added rows: " + addedMap.size());
//                XSSFSheet addSheet2 = book.createSheet("Added2");
//                outOneDiffSheet(addSheet2, addedMap);
            }
            else {
                System.out.println("There are no added rows.");
            }

            if (changedMap.size() > 0) {
                XSSFSheet changedSheet = book.createSheet("Changed");
                copySheetWithFilter(newSheet, changedSheet, changedMap, null, maxColumn);
                System.out.println("* Changed rows: " + changedMap.size());
            }
            else {
                System.out.println("There are no changed rows.");
            }

            if (matchedMap.size() > 0) { // Rows matched with MxWeb requirements
                if (maxColumn > Requirement.MXWEB_RELEASE) {
                    XSSFSheet matchedSheet = book.createSheet("MX Matched");
                    XLSUtil.copySheet(newFileName, matchedSheet, maxColumn);
                    System.out.println("= MxWeb matched rows: " + matchedMap.size());
                    markOneSheet(matchedSheet, matchedMap, null,false);
                    // Output matched results
                    int maxRow = matchedSheet.getLastRowNum();
                    for (int i = 0; i < maxRow; i++) {
                        XSSFRow row = matchedSheet.getRow(i);
                        for (Map.Entry<String, Requirement> item : matchedMap.entrySet()) {
                            if (item.getValue().getRow() == i) {
                                String mxWebRelease = item.getValue().getSource_req();
                                if (row.getLastCellNum() > Requirement.MXWEB_RELEASE) {
                                    row.getCell(16).setCellValue(mxWebRelease);
                                } else {
                                    XSSFCell cell = row.createCell(Requirement.MXWEB_RELEASE);
                                    cell.setCellValue(mxWebRelease);
                                }
                            }
                        }
                    }
                    moveValuesToParents(matchedSheet, 16);
                }
            }
            else {
                System.out.println("There are no matched rows with MxWeb.");
            }

            if (missedMap.size() > 0) { // MX requirenets rows missed in our requirement
                XSSFSheet missedSheet = book.createSheet("MX Missed");
                MxRequirement.outMissedSheet(missedSheet, missedMap);
                System.out.println("X MxWeb missed rows: " + missedMap.size());
            }
            else {
                System.out.println("There are no missed rows with MxWeb.");
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
     * Mark rows in specified sheet according requirements in specified map. Mark style depends on isDeleted flag
     * @param sheet - sheet to mark
     * @param array - array with specified rows (deleted, added or changed)
     * @param columnDetails - array with column indexes have to be marked (required for changed rows only)
     * @param isDeleted - when true rows mark as grey strikeout (deleted), when false - mark as red (inserted), when null - mark only columns in columnDetails
     */
    private static void markOneSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array,
                                     LinkedHashMap<String, List<Integer>> columnDetails, Boolean isDeleted) {

        int lastRow = sheet.getLastRowNum();
        Requirement req = new Requirement();

        // Common styles for all group levels (with outline levels)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        for (int i = 0; i <= lastRow; i++) {

            if (i <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            XSSFRow row = sheet.getRow(i);
            req.loadFromRow(row);

            ArrayList<XSSFCellStyle> styles = new ArrayList<>(); // Styles for current row

            List<Integer> markedColumns = null; // List of marked columns (to mark changes only specified cells)

            int outlineLevel = row.getOutlineLevel();

            if (array.containsKey(req.id)) {

                if (columnDetails != null) {
                    markedColumns = columnDetails.get(req.id);
                    if (markedColumns == null) {
                        System.out.println("ERROR: arrays mismatch - can't obtain changes details for row " + i);
                        continue;
                    }
                }

                if (!groupStyles.containsKey(outlineLevel)) {
                    // Styles for row with unknown outline level
                    // Copy style from existing cell and correct: all styles after specified row are common - takes it from array

                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        XSSFCellStyle newCellStyle = sheet.getWorkbook().createCellStyle();
                        newCellStyle.cloneStyleFrom(cell.getCellStyle());
                        Font font = sheet.getWorkbook().createFont();
                        if (isDeleted != null) {
                            if (isDeleted) {
                                font.setBold(true);
                                font.setStrikeout(true);
                                font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
                            } else {
                                font.setBold(true);
                                font.setColor(IndexedColors.RED.getIndex());
                            }
                        }
                        else { // Mark changed rows (only specified columns)
                            font.setBold(true);
                            font.setColor(IndexedColors.BLUE.getIndex());
                        }
                        newCellStyle.setFont(font);
                        styles.add(newCellStyle);
                    }
                    groupStyles.put(outlineLevel, styles);
                }

                if (groupStyles.containsKey(outlineLevel)) {
                    styles = groupStyles.get(outlineLevel);
                    if (styles != null) {
                        if (markedColumns == null) { // Mark all cells in one row
                            for (int j = 0; j < row.getLastCellNum(); j++) {
                                XSSFCell cell = row.getCell(j);
                                if (j < styles.size()) cell.setCellStyle(styles.get(j));
                            }
                        }
                        else { // Mark specified cells only
                            for (int j = 0; j < markedColumns.size(); j++) {
                                int k = markedColumns.get(j);
                                XSSFCell cell = row.getCell(k);
                                if (cell != null) {
                                    if (k < styles.size()) cell.setCellStyle(styles.get(k));
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * Copies rows that contain in specified filter map and copies all its parent rows to specified sheet
     * @param sourceSheet - source sheet
     * @param targetSheet - source sheet
     * @param filterMap - array with specified rows
     * @param isDeleted - when true rows mark as grey strikeout (deleted), when false - mark as red (inserted), when null - do nothing with rows style
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     */
    private static void copySheetWithFilter(XSSFSheet sourceSheet, XSSFSheet targetSheet,
                                        LinkedHashMap<String, Requirement> filterMap,
                                        Boolean isDeleted, int maxColumn) {

        TreeMap<Integer, Requirement> parents = new TreeMap<>(); // Rows for parent nodes (outline levels and requirements)
        TreeMap<Integer, Requirement> prevParents = new TreeMap<>(); // Parents for previous row

        // Common styles for suitable rows with all outline levels
        HashMap<Integer, ArrayList<XSSFCellStyle>> rowStyles = new HashMap<>();
        // Styles for parent rows with all outline levels (will be corrected to GREY)
        HashMap<Integer, ArrayList<XSSFCellStyle>> parentStyles = new HashMap<>();

        int newRowNum = 0; // Current row number in target sheet

        int lastRow = sourceSheet.getLastRowNum();

        for (int i = 0; i <= lastRow; i++) { // View all source sheet

            if (i <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            Requirement req = new Requirement(); // Current requirement
            XSSFRow row = sourceSheet.getRow(i);
            req.loadFromRow(row);

            if (filterMap.containsKey(req.id)) { // Suitable row?

                int outlineLevel = row.getOutlineLevel();

                if (!rowStyles.containsKey(outlineLevel)) { // Fill styles map for suitable rows
                    ArrayList<XSSFCellStyle> styles = new ArrayList<>();
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        XSSFCellStyle newCellStyle = targetSheet.getWorkbook().createCellStyle();
                        newCellStyle.cloneStyleFrom(cell.getCellStyle());
                        Font font = targetSheet.getWorkbook().createFont();
                        if (isDeleted != null) {
                            if (isDeleted) {
                                font.setBold(true);
                                font.setStrikeout(true);
                            } else {
                                font.setBold(true);
                                font.setColor(IndexedColors.RED.getIndex());
                            }
                        }
                        newCellStyle.setFont(font);
                        styles.add(newCellStyle);
                    }
                    rowStyles.put(outlineLevel, styles);
                }

                // Get all parent nodes for one row suitable for filter
                int parentRowNum = i;
                parents.clear();
                for (int j = outlineLevel - 1; j >= 0; j--) {
                    while (true) {
                        parentRowNum--;
                        if (parentRowNum < 0) {
                            System.out.println("ERROR. Can't find full path for row " + (row.getRowNum() + 1));
                            break;
                        }
                        XSSFRow prevRow = row.getSheet().getRow(parentRowNum);
                        int prevOutlineLevel = prevRow.getOutlineLevel();
                        if (prevOutlineLevel == j) {
                            Requirement parentReq = new Requirement(); // One of parent requirement (for iterations by parents)
                            parentReq.loadFromRow(prevRow);
                            parents.put(prevOutlineLevel, parentReq);
                            break;
                        } else if (prevOutlineLevel < j) {
                            System.out.println("ERROR. Outline levels sequence violation for row " + (row.getRowNum() + 1));
                            break;
                        }
                    }
                }

                boolean isSamePath = false;
                isSamePath = (prevParents.size() == parents.size());
                if (isSamePath) {
                    for (Map.Entry<Integer, Requirement> item : parents.entrySet()) {
                        int key = item.getKey();
                        if (!(item.getValue().id.equals(prevParents.get(key).id))) {
                            isSamePath = false;
                            break;
                        }
                    }
                }

                if (!isSamePath) {
                    // Copy all parent rows and fill styles map for parent rows
                    for (Map.Entry<Integer, Requirement> item : parents.entrySet()) {
                        Requirement parentReq = item.getValue();
                        parentRowNum = parentReq.getRow();
                        XSSFRow parentRow = sourceSheet.getRow(parentRowNum);
                        int parentOutlineLevel = parentRow.getOutlineLevel();
                        if (!parentStyles.containsKey(parentOutlineLevel)) {
                            ArrayList<XSSFCellStyle> styles = new ArrayList<>();
                            // Copy style from existing cell and correct: all styles after specified row are common - takes it from array
                            for (int j = 0; j < parentRow.getLastCellNum(); j++) {
                                XSSFCell cell = parentRow.getCell(j);
                                XSSFCellStyle newCellStyle = targetSheet.getWorkbook().createCellStyle();
                                newCellStyle.cloneStyleFrom(cell.getCellStyle());
                                Font font = targetSheet.getWorkbook().createFont();
                                font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
                                newCellStyle.setFont(font);
                                styles.add(newCellStyle);
                            }
                            parentStyles.put(parentOutlineLevel, styles);
                        }
                        XLSUtil.copyRow(sourceSheet, targetSheet, parentRowNum, newRowNum, maxColumn, parentStyles.get(parentOutlineLevel));
                        newRowNum++;
                    }
                }
                // Copy main row suitable for filter
                XLSUtil.copyRow(sourceSheet, targetSheet, i, newRowNum, maxColumn, rowStyles.get(outlineLevel));
                prevParents = (TreeMap<Integer, Requirement>) parents.clone();
                // prevParents.put(outlineLevel, req);
                newRowNum++;
            }
        }
    }

    /**
     * Outputs one Excel sheet with added or deleted rows with full path rows above
     * @param sheet - Excel sheet
     * @param array - array to fill sheet
     */
    private static void outOneDiffSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array) {
        XSSFCellStyle style = sheet.getWorkbook().createCellStyle();
        XSSFFont font = sheet.getWorkbook().createFont();
        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex()); // Color for path rows
        style.setFont(font);

        int rowNum = 0;
        int level = 0;
        String oldPathStr = "";
//        String[] oldPathArr = {};
        for (Map.Entry<String, Requirement> item : array.entrySet()) {
            String pathStr = getOnlyPath(item.getKey());
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
//                oldPathArr = pathArr;
            }
            XSSFRow row = sheet.createRow(rowNum);
            XSSFCell path = row.createCell(0);
            path.setCellValue(level + 1);
            XSSFCell name = row.createCell(1);
            name.setCellValue(getOnlyName(item.getKey()));
            rowNum++;
        }
    }

    /**
     * Returns name from full path (path with name), divided by "|"
     * @param item - path with name
     * @return - only name
     */
    private static String getOnlyName(String item) {
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
    private static String getOnlyPath(String item) {
        String name = "";
        if (item != null) {
            int pos = item.lastIndexOf('|');
            if (pos >= 0) name = item.substring(0, pos);
        }
        return name;
    }

    /**
     * Copies values in specified column to all parent rows. If rows for one parent contains different values it merges in list separated by "\n"
     * @param sheet - sheet
     * @param column - column index (starts from 0)
     */
    private static void moveValuesToParents(XSSFSheet sheet, int column) {
        // Get all parent nodes for one row suitable for filter
        int lastRow = sheet.getLastRowNum();
        LinkedHashMap<Integer, XSSFRow> parents = new LinkedHashMap<>();

        for (int i = 0; i <= lastRow; i++) { // View all sheet

            XSSFRow row = sheet.getRow(i);
            if (row.getLastCellNum() < column) continue;
            String value = row.getCell(column).getStringCellValue();
            if (value == null || value.isEmpty()) continue;
            int outlineLevel = row.getOutlineLevel();

            int parentRowNum = i;
            parents.clear();
            for (int j = outlineLevel - 1; j >= 0; j--) {
                while (true) {
                    parentRowNum--;
                    if (parentRowNum < 0) {
                        System.out.println("ERROR. Can't find full path for row " + (row.getRowNum() + 1));
                        break;
                    }
                    XSSFRow prevRow = row.getSheet().getRow(parentRowNum);
                    int prevOutlineLevel = prevRow.getOutlineLevel();
                    if (prevOutlineLevel == j) {
                        String parentValue = prevRow.getCell(column).getStringCellValue();
                        if (parentValue != null && !parentValue.isEmpty()) {
                            String values[] = parentValue.split("\n");
                            ArrayList<String> set = new ArrayList<>();
                            Collections.addAll(set, values);
                            boolean isFound = set.contains(value);
                            if (!isFound) {
                                set.add(value);
                                Collections.sort(set);
                                StringBuilder cellValue = new StringBuilder();
                                boolean isFirst = true;
                                for (String s : set) {
                                    if (isFirst) {
                                        isFirst = false;
                                    }
                                    else {
                                        cellValue.append("\n");
                                    }
                                    cellValue.append(s);
                                }
                                prevRow.getCell(column).setCellValue(cellValue.toString());
                            }
                        }
                        else {
                            prevRow.getCell(column).setCellValue(value);
                        }
                        break;
                    } else if (prevOutlineLevel < j) {
                        System.out.println("ERROR. Outline levels sequence violation for row " + (row.getRowNum() + 1));
                        break;
                    }
                }
            }
        }
    }
}
