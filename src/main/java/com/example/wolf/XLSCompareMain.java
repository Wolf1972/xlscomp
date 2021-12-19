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
        String mergeMode = "e";

        Options options = new Options();
        options.addOption("o", "old", true, "Old XLSX file for compare");
        options.addOption("n", "new", true, "New XLSX file for compare");
        options.addOption("r", "result", true, "XLSX file for result");
        options.addOption("m", "merge", true, "[Merge cells for common rows from old file to new (-\"ma\" - all different, -\"me\" - if only new cell is empty)]");
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

            System.out.println("Creating results book...");
            XSSFWorkbook book = new XSSFWorkbook();

            XSSFSheet currentSheet = book.createSheet("Current");
            XLSUtil.copySheet(newName, currentSheet, maxColumn);
            XLSUtil.groupSheet(currentSheet, Requirement.HEADER_LAST_ROW + 1);

            XSSFSheet oldSheet = book.createSheet("Old");
            XLSUtil.copySheet(oldName, oldSheet, maxColumn);
            XLSUtil.groupSheet(oldSheet, Requirement.HEADER_LAST_ROW + 1);

            XSSFSheet newSheet = book.createSheet("New");
            XLSUtil.copySheet(newName, newSheet, maxColumn);
            XLSUtil.groupSheet(newSheet, Requirement.HEADER_LAST_ROW + 1);

            System.out.println("Results book created.");

            LinkedHashMap<String, Requirement> deletedMap = new LinkedHashMap<>();
            LinkedHashMap<String, Requirement> addedMap = new LinkedHashMap<>();

            LinkedHashMap<String, Requirement> changedMap = new LinkedHashMap<>(); // Map with only changed rows
            LinkedHashMap<String, List<Integer>> changedMapDetails = new LinkedHashMap<>(); // Map with columns has changed for every row has changed

            if (mergeMode != null && !mergeMode.isEmpty()) {
                // Copy old sheet to both sheets: Current and New (they are same)
                mergeSheet(oldSheet, currentSheet, oldMap, newMap, !mergeMode.contains("a"), maxColumn, false);
                mergeSheet(oldSheet, newSheet, oldMap, newMap, !mergeMode.contains("a"), maxColumn, true);
            }

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

            System.out.println("Requirements compared.");

            LinkedHashMap<String, Requirement> matchedMap = new LinkedHashMap<>(); // Rows from mxWeb were matched without errors
            LinkedHashMap<Integer, MxRequirement> missedMap = new LinkedHashMap<>(); // Rows from MxWeb requirements were missed when matching

            if (mxWebName != null) {

                fileName = mxWebName;
                System.out.println("Loading mxWeb file " + fileName);
                LinkedHashMap<Integer, MxRequirement> mxwebMap = MxRequirement.readFromExcel(fileName);
                System.out.println("Rows loaded: " + mxwebMap.size());

                matchedMap = mxWebCheck(newMap, mxwebMap); // Direct check - search for MxWeb requirements were specified in our requirements

                missedMap = mxWebMissing(newMap, mxwebMap); // Reverse check - try to find MxWeb requirements were missed in our requirements

                System.out.println("MxWeb requirements checked.");
            }

            outResult(book, oldName, newName, maxColumn,
                      addedMap, deletedMap,
                      changedMap, changedMapDetails,
                      matchedMap, missedMap);

            try {
                book.write(new FileOutputStream(resultName));
                book.close();
            }
            catch (IOException e) {
                System.out.println("ERROR when saving results file: " + resultName);
            }

        }
        catch (IOException e) {
            System.out.println("XLSX file read error: " + fileName);
        }

    }

    /**
     * Outputs compare results
     * @param book - results book
     * @param oldFileName - old file name
     * @param newFileName - new file name for compare
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     * @param addedMap - map with added rows
     * @param deletedMap - map with deleted rows
     * @param changedMap - map with changed rows
     * @param changedDetailsMap - map with changed details
     * @param matchedMap - map with matched with MaxWeb requirement rows
     * @param missedMap - map with MxWeb requirements rows were missed when matching
     */
    private static void outResult(XSSFWorkbook book, String oldFileName, String newFileName, int maxColumn,
                                  LinkedHashMap<String, Requirement> addedMap, LinkedHashMap<String, Requirement> deletedMap,
                                  LinkedHashMap<String, Requirement> changedMap, LinkedHashMap<String, List<Integer>> changedDetailsMap,
                                  LinkedHashMap<String, Requirement> matchedMap, LinkedHashMap<Integer, MxRequirement> missedMap) {

        System.out.println("Output results.");

        if (addedMap.size() > 0 || deletedMap.size() > 0) {

            XSSFSheet currentSheet = book.getSheet("Current");

            XSSFSheet oldSheet = book.getSheet("Old");
            if (deletedMap.size() > 0) markOneSheet(oldSheet, deletedMap, null, MarkRowType.DELETED);
            if (changedMap.size() > 0) markOneSheet(oldSheet, changedMap, changedDetailsMap, MarkRowType.CHANGED);

            XSSFSheet newSheet = book.getSheet("New");
            if (addedMap.size() > 0) markOneSheet(newSheet, addedMap, null, MarkRowType.ADDED);
            if (changedMap.size() > 0) markOneSheet(newSheet, changedMap, changedDetailsMap, MarkRowType.CHANGED);

            if (deletedMap.size() > 0) {
                XSSFSheet delSheet = book.createSheet("Deleted");
                XLSUtil.copyHeader(oldSheet, delSheet, maxColumn);
                copySheetWithFilter(oldSheet, delSheet, deletedMap, maxColumn);
                XLSUtil.groupSheet(delSheet, Requirement.HEADER_LAST_ROW + 1);
                System.out.println("- Deleted rows: " + deletedMap.size());
            }
            else {
                System.out.println("- There are no deleted rows.");
            }

            if (addedMap.size() > 0) {
                XSSFSheet addSheet = book.createSheet("Added");
                XLSUtil.copyHeader(oldSheet, addSheet, maxColumn);
                copySheetWithFilter(newSheet, addSheet, addedMap, maxColumn);
                XLSUtil.groupSheet(addSheet, Requirement.HEADER_LAST_ROW + 1);
                System.out.println("+ Added rows: " + addedMap.size());
            }
            else {
                System.out.println("+ There are no added rows.");
            }

            if (changedMap.size() > 0) {
                XSSFSheet changedSheet = book.createSheet("Changed");
                XLSUtil.copyHeader(newSheet, changedSheet, maxColumn);
                copySheetWithFilter(newSheet, changedSheet, changedMap, maxColumn);
                XLSUtil.groupSheet(changedSheet, Requirement.HEADER_LAST_ROW + 1);
                markOneSheet(newSheet, changedMap, changedDetailsMap, MarkRowType.CHANGED);
                markOneSheet(changedSheet, changedMap, changedDetailsMap, MarkRowType.CHANGED);
                System.out.println("* Changed rows: " + changedMap.size());
            }
            else {
                System.out.println("* There are no changed rows.");
            }

            if (matchedMap.size() > 0) { // Rows matched with MxWeb requirements
                if (maxColumn > Requirement.MXWEB_RELEASE) {
                    XSSFSheet matchedSheet = book.createSheet("MX Matched");
                    XLSUtil.copySheet(newFileName, matchedSheet, maxColumn);
                    XLSUtil.groupSheet(matchedSheet, Requirement.HEADER_LAST_ROW + 1);
                    markOneSheet(matchedSheet, matchedMap, null,MarkRowType.CHANGED);
                    System.out.println("= MxWeb matched rows: " + matchedMap.size());
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
                System.out.println("= There are no matched rows with MxWeb.");
            }

            if (missedMap.size() > 0) { // MX requirenets rows missed in our requirement
                XSSFSheet missedSheet = book.createSheet("MX Missed");
                MxRequirement.outMissedSheet(missedSheet, missedMap);
                System.out.println("X MxWeb missed rows: " + missedMap.size());
            }
            else {
                System.out.println("X There are no missed rows with MxWeb.");
            }
        }
    }

    /**
     * Deletes all rows from sheet according requirements in specified map
     * @param sheet - sheet to delete rows
     * @param array - array with specified rows to delete
     */
    private static void deleteRows(XSSFSheet sheet, LinkedHashMap<String, Requirement> array) {

        int lastRow = sheet.getLastRowNum();
        Requirement req = new Requirement();

        for (int i = lastRow; i >= 0; i--) {

            if (i <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            XSSFRow row = sheet.getRow(i);
            req.loadFromRow(row);

            if (array.containsKey(req.id)) {
                sheet.removeRow(row);
            }
        }
    }

    /**
     * Function copies cells from old sheet to new sheet with specified mode: all different cells or only empty cells
     * @param newSheet - sheet with old requirements
     * @param oldSheet - sheet with new requirements
     * @param oldMap - old requirement
     * @param newMap - new requirement
     * @param isEmptyOnly - copy mode (when true: only empty cells in new sheet, false: all cells than have difference with old cells)
     * @param maxColumn - max number of columns
     * @param isNewRefresh - when true: refresh newMap with requirements that have been copied
     * @return - map with requirements were copied
     */
    private static LinkedHashMap<String, Requirement> mergeSheet(XSSFSheet oldSheet, XSSFSheet newSheet,
                                                                 LinkedHashMap <String, Requirement> oldMap,
                                                                 LinkedHashMap <String, Requirement> newMap,
                                                                 boolean isEmptyOnly,
                                                                 int maxColumn,
                                                                 boolean isNewRefresh) {

        System.out.println("Copying requirements attributes from " + oldSheet.getSheetName() + " to " + newSheet.getSheetName() + " " + (isEmptyOnly ? "(only empty cells)":"(all different cells)" + "..."));

        LinkedHashMap<String, Requirement> copiedMap = new LinkedHashMap<>();

        for (Map.Entry<String, Requirement> item : oldMap.entrySet()) {
            String key = item.getKey();
            Requirement req = item.getValue();
            if (newMap.containsKey(key)) {
                Requirement newReq = newMap.get(key);
                if (newReq != null) {
                    List<Integer> changedColumns = newReq.compare(req);
                    if (changedColumns != null && changedColumns.size() > 0) {
                        copiedMap.put(key, req);
//                        System.out.print("Copy row " + req.getRow() + " to row " + newReq.getRow() + ", columns " + changedColumns + "... ");
                        XLSUtil.copyRow(oldSheet, newSheet, false, req.getRow(), newReq.getRow(), maxColumn, changedColumns, isEmptyOnly, null);
                        XSSFRow newRow = newSheet.getRow(newReq.getRow());
                        if (isNewRefresh) newReq.loadFromRow(newRow); // Refresh new requirement with copying results
//                        System.out.println("Done.");
                    }
                }
            }
        }
        System.out.println("Copying done. " + copiedMap.size() + " row(s) copied.");
        return copiedMap;
    }

    /**
     * Mark rows in specified sheet according requirements in specified map. Mark style depends on specified type
     * @param sheet - sheet to mark
     * @param array - array with specified rows (deleted, added or changed)
     * @param columnDetails - array with column indexes have to be marked (required for changed rows only)
     * @param mark - mark type (deleted, inserted or changed)
     */
    private static void markOneSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array,
                                     LinkedHashMap<String, List<Integer>> columnDetails, MarkRowType mark) {

        int lastRow = sheet.getLastRowNum();
        Requirement req = new Requirement();

        // Common styles for all group levels (with outline levels)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        for (int i = Requirement.HEADER_LAST_ROW + 1; i <= lastRow; i++) {

            XSSFRow row = sheet.getRow(i);
            req.loadFromRow(row);

            ArrayList<XSSFCellStyle> styles = new ArrayList<>(); // Styles for current row

            List<Integer> markedColumns = null; // List of marked columns (to mark changes only specified cells)

            int outlineLevel = row.getOutlineLevel();

            if (array.containsKey(req.id)) {

                if (columnDetails != null) { // When change mode we have to check according between changed rows and changed details (columns) for each row
                    markedColumns = columnDetails.get(req.id);
                    if (markedColumns == null) {
                        System.out.println("ERROR: arrays mismatch - can't obtain changes details for row " + i);
                        continue;
                    }
                }

                if (!groupStyles.containsKey(outlineLevel)) { // Styles for row with new outline level
                    styles = XLSUtil.modifyCellStyles(sheet, row, mark);
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
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     */
    private static void copySheetWithFilter(XSSFSheet sourceSheet, XSSFSheet targetSheet,
                                            LinkedHashMap<String, Requirement> filterMap,
                                            int maxColumn) {

        TreeMap<Integer, Requirement> parents = new TreeMap<>(); // Rows for parent nodes (outline levels and requirements)
        TreeMap<Integer, Requirement> prevParents = new TreeMap<>(); // Parents for previous row

        // Common styles for suitable rows with all outline levels
        HashMap<Integer, ArrayList<XSSFCellStyle>> rowStyles = new HashMap<>();
        // Styles for parent rows with all outline levels
        HashMap<Integer, ArrayList<XSSFCellStyle>> parentStyles = new HashMap<>();

        int newRowNum = Requirement.HEADER_LAST_ROW + 1; // Current row number in target sheet
        String oldName = null; // Previous name was copied

        int lastRow = sourceSheet.getLastRowNum();

        for (int i = Requirement.HEADER_LAST_ROW + 1; i <= lastRow; i++) { // View all source sheet

            Requirement req = new Requirement(); // Current requirement
            XSSFRow row = sourceSheet.getRow(i);
            req.loadFromRow(row);

            if (filterMap.containsKey(req.id)) { // Suitable row?

                int outlineLevel = row.getOutlineLevel();

                if (!rowStyles.containsKey(outlineLevel)) { // Fill styles map for suitable rows
                    ArrayList<XSSFCellStyle> styles = XLSUtil.getRowStyles(sourceSheet, i, maxColumn);
                    styles = XLSUtil.cloneRowStyles(targetSheet, styles);
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

                if (!isSamePath) { // Outline level has been changed for current row
                    // Fill styles map for parent rows
                    for (Map.Entry<Integer, Requirement> item : parents.entrySet()) {
                        Requirement parentReq = item.getValue();
                        parentRowNum = parentReq.getRow();
                        XSSFRow parentRow = sourceSheet.getRow(parentRowNum);
                        int parentOutlineLevel = parentRow.getOutlineLevel();
                        if (!parentStyles.containsKey(parentOutlineLevel)) {
                            ArrayList<XSSFCellStyle> styles = XLSUtil.getRowStyles(sourceSheet, parentRowNum, maxColumn);
                            XLSUtil.cloneRowStyles(targetSheet, styles);
                            styles = XLSUtil.modifyCellStyles(targetSheet, parentRow, MarkRowType.PARENT);
                            parentStyles.put(parentOutlineLevel, styles);
                        }
                        // Copy only parent rows that mismatch with previous parent rows
                        if (prevParents.size() > 0 && prevParents.containsValue(item.getValue())) {
                            // One of parent rows that already was copied
                        }
                        else if (oldName == null || !oldName.equals(item.getValue().getName())) {
                            XLSUtil.copyRow(sourceSheet, targetSheet, true, parentRowNum, newRowNum,
                                            maxColumn, null, false, parentStyles.get(parentOutlineLevel));
                            newRowNum++;
                        }
                    }
                }
                // Copy main row suitable for filter
                XLSUtil.copyRow(sourceSheet, targetSheet, true, i, newRowNum,
                                 maxColumn, null, false, rowStyles.get(outlineLevel));
                prevParents = (TreeMap<Integer, Requirement>) parents.clone();
                oldName = req.getName();
                // prevParents.put(outlineLevel, req);
                newRowNum++;
            }
        }
    }

    /**
     * Functions compares new sheet with requirements with MxWeb requirements and returns array with matched
     * @param checkMap - map with our requirements (new requirements)
     * @param mxWebMap - map with MxWeb requirements
     * @return - map with our requirements matched with MxWeb requirements
     */
    private static LinkedHashMap<String, Requirement> mxWebCheck(HashMap<String, Requirement> checkMap,
                                                           HashMap<Integer, MxRequirement> mxWebMap) {

        LinkedHashMap<String, Requirement> matchedMap = new LinkedHashMap<>();

        System.out.println("Trying to match our requirements with MxWeb requirements...");

        for (Map.Entry<String, Requirement> item : checkMap.entrySet()) {
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
                    for (Map.Entry<Integer, MxRequirement> mxitem : mxWebMap.entrySet()) {
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
        return matchedMap;
    }

    /**
     * Function preforms reverse check - looking for MxWeb requirements were missing in our requirements
     * @param checkMap - map with our requirements
     * @param mxWebMap - map with MxWeb requirements
     * @return - map with MxWeb requirements that were missing in our requirements
     */
    private static LinkedHashMap<Integer, MxRequirement> mxWebMissing(HashMap<String, Requirement> checkMap,
                                                                HashMap<Integer, MxRequirement> mxWebMap) {

        LinkedHashMap <Integer, MxRequirement> missedMap = new LinkedHashMap<>();

        System.out.println("Trying to match MX requirements with our requirements...");
        // Reverse check - try to find MxWeb requirement has missed in our sheet
        for (Map.Entry<Integer, MxRequirement> mxitem : mxWebMap.entrySet()) {

            String searchFor = mxitem.getValue().getMxwebid();
            String release = mxitem.getValue().getRelease();
            if (searchFor != null && !searchFor.isEmpty()) {
                if (release == null || release.isEmpty()) {
                    System.out.println("WARN. MxWeb requirement " + searchFor + " without release specified.");
                }

                boolean isFound = false;
                for (Map.Entry<String, Requirement> item : checkMap.entrySet()) {
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
        return missedMap;
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
