package com.example.wolf;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.*;

public class XLSCompareMain {

    public static void main (String[] args) {
        System.out.println("Trying to compare XLSX with requirements hierarchy.");

        String dir = System.getProperty("user.dir");
        String oldName = dir + "\\data\\old.xlsx";
        String newName = dir + "\\data\\new.xlsx";
        String resultName = dir + "\\data\\result.xlsx";
        String mxwebName = dir + "\\data\\mxweb.xlsx";

        if (args.length >= 4) {
            oldName = args[0];
            newName = args[1];
            mxwebName = args[2];
            resultName = args[3];
        }
        else if (args.length >= 3) {
            oldName = args[0];
            newName = args[1];
            resultName = args[2];
        }
        else {
            System.out.println("Usage: <old file> <new file to compare> <file with results>");
            System.out.println("   or: <old file> <new file to compare> <mxweb file> <file with results>");
            System.out.println("You didn't define any parameter, try to default names...");
            System.out.println("Current directory: " + dir);
        }
        System.out.println();
        doCompare(oldName, newName, mxwebName, resultName);
        System.out.println("Done.");
    }

    /**
     * Compare requirements procedure
     * @param oldName - old file name
     * @param newName - new file name for compare
     * @param mxwebName - file name with mxweb requirements
     * @param resultName - file name for result
     */
    private static void doCompare(String oldName, String newName, String mxwebName, String resultName) {

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

            for (Map.Entry<String, Requirement> item : oldMap.entrySet()) {
                String key = item.getKey();
                Requirement req = item.getValue();
                if (!newMap.containsKey(key)) {
                    deletedMap.put(key, req);
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

            LinkedHashMap<String, Requirement> mergedMap = new LinkedHashMap<>(); // Rows were merged without errors
            LinkedHashMap<Integer, MxRequirement> missedMap = new LinkedHashMap<>(); // Rows with MxWeb requirements were missed when merge

            if (mxwebName != null) {

                System.out.println("Trying to merge...");

                fileName = mxwebName;
                System.out.println("Loading mxweb file " + fileName);
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
                                    mergedMap.put(item.getKey(), item.getValue());
                                    isFound = true;
                                    break;
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
                                        mergedMap.put(item.getKey(), item.getValue());
                                        String oldRelease = item.getValue().getSource_req();
                                        if (oldRelease != null && !oldRelease.isEmpty() && !oldRelease.equals(release)) {
                                            System.out.println("ERROR. Row " + (item.getValue().getRow() + 1) + " already has release " + oldRelease);
                                        }
                                        else {
                                            item.getValue().setSource_req(release);
                                        }
                                        isFound = true;
                                        break;
                                    }
                                }
                                if (isFound) break;
                            }
                        }
                        if (!isFound) {
                            System.out.println("ERROR. MxWeb requirement " + (mxitem.getValue().getMxwebid()) + " has not found in our sheet.");
                            missedMap.put(mxitem.getKey(), mxitem.getValue());
                        }
                    }
                }

                System.out.println("Merged.");
            }

            outResult(oldName, newName, resultName, addedMap, deletedMap, mergedMap, missedMap);

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
     * @param addedMap - map with added rows
     * @param deletedMap - map with deleted rows
     * @param mergedMap - map with merged rows
     * @param missedMap - map with MxWeb requirements rows were missed when merge
     */
    private static void outResult(String oldFileName, String newFileName, String resultFileName,
                                  LinkedHashMap<String, Requirement> addedMap, LinkedHashMap<String, Requirement> deletedMap,
                                  LinkedHashMap<String, Requirement> mergedMap, LinkedHashMap<Integer, MxRequirement> missedMap) {

        XSSFWorkbook book = new XSSFWorkbook();

        if (addedMap.size() > 0 || deletedMap.size() > 0) {

            XSSFSheet currentSheet = book.createSheet("Current");
            XLSUtil.copySheet(newFileName, currentSheet);

            XSSFSheet oldSheet = book.createSheet("Old");
            XLSUtil.copySheet(oldFileName, oldSheet);
            if (deletedMap.size() > 0) {
                System.out.println("- Deleted rows: " + deletedMap.size());
                markOneSheet(oldSheet, deletedMap, true);
                XSSFSheet delSheet = book.createSheet("Deleted");
                copySheetFilter(oldSheet, delSheet, deletedMap, true);
                XSSFSheet delSheet2 = book.createSheet("Deleted2");
                outOneDiffSheet(delSheet2, deletedMap);
            }
            else {
                System.out.println("There are no deleted rows.");
            }

            XSSFSheet newSheet = book.createSheet("New");
            XLSUtil.copySheet(newFileName, newSheet);
            if (addedMap.size() > 0) {
                System.out.println("+ Added rows: " + addedMap.size());
                markOneSheet(newSheet, addedMap, false);
                XSSFSheet addSheet = book.createSheet("Added");
                copySheetFilter(newSheet, addSheet, addedMap, false);
                XSSFSheet addSheet2 = book.createSheet("Added2");
                outOneDiffSheet(addSheet2, addedMap);
            }
            else {
                System.out.println("There are no added rows.");
            }

            if (mergedMap.size() > 0) { // TODO
                XSSFSheet mergedSheet = book.createSheet("Merged");
                XLSUtil.copySheet(newFileName, mergedSheet);
                System.out.println("= Merged rows: " + mergedMap.size());
                markOneSheet(mergedSheet, mergedMap, false);
                // Output merged results
                int maxRow = mergedSheet.getLastRowNum();
                for (int i = 0; i < maxRow; i++) {
                    XSSFRow row = mergedSheet.getRow(i);
                    for (Map.Entry<String, Requirement> item : mergedMap.entrySet()) {
                        if (item.getValue().getRow() == i) {
                            String release = item.getValue().getSource_req();
                            if (row.getLastCellNum() > 16) {
                                if (!item.getValue().getSource_req().equals(release))
                                    row.getCell(16).setCellValue(release);
                            } else {
                                XSSFCell cell = row.createCell(16);
                                cell.setCellValue(release);
                            }
                        }
                    }
                }
            }
            else {
                System.out.println("There are no merged rows.");
            }

            if (missedMap.size() > 0) {
                XSSFSheet missedSheet = book.createSheet("Missed");
                MxRequirement.outMissedSheet(missedSheet, missedMap);
                System.out.println("X Missed rows: " + missedMap.size());
            }
            else {
                System.out.println("There are no missed rows.");
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
     * @param array - array with specified rows
     * @param isDeleted - when true rows mark as grey strikeout (deleted), when false - mark as red (inserted)
     */
    private static void markOneSheet(XSSFSheet sheet, LinkedHashMap<String, Requirement> array, boolean isDeleted) {

        int lastRow = sheet.getLastRowNum();
        Requirement req = new Requirement();

        // Common styles for all group levels (with outline levels)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        for (int i = 0; i <= lastRow; i++) {

            if (i <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            XSSFRow row = sheet.getRow(i);
            req.loadFromRow(row);

            ArrayList<XSSFCellStyle> styles = new ArrayList<>(); // Styles for current row

            int outlineLevel = row.getOutlineLevel();

            if (array.containsKey(req.id)) {

                if (!groupStyles.containsKey(outlineLevel)) {
                    // Styles for row with unknown outline level
                    // Copy style from existing cell and correct: all styles after specified row are common - takes it from array
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        XSSFCellStyle newCellStyle = sheet.getWorkbook().createCellStyle();
                        newCellStyle.cloneStyleFrom(cell.getCellStyle());
                        Font font = sheet.getWorkbook().createFont();
                        if (isDeleted) {
                            font.setBold(true);
                            font.setStrikeout(true);
                            font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
                        }
                        else {
                            font.setBold(true);
                            font.setColor(IndexedColors.RED.getIndex());
                        }
                        newCellStyle.setFont(font);
                        styles.add(newCellStyle);
                    }
                    groupStyles.put(outlineLevel, styles);
                }

                if (groupStyles.containsKey(outlineLevel)) {
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        styles = groupStyles.get(outlineLevel);
                        if (j < styles.size()) cell.setCellStyle(styles.get(j));
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
     * @param isDeleted - when true rows mark as grey strikeout (deleted), when false - mark as red (inserted)
     */
    private static void copySheetFilter(XSSFSheet sourceSheet, XSSFSheet targetSheet, LinkedHashMap<String, Requirement> filterMap, boolean isDeleted) {

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
                        if (isDeleted) {
                            font.setBold(true);
                            font.setStrikeout(true);
                        }
                        else {
                            font.setBold(true);
                            font.setColor(IndexedColors.RED.getIndex());
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
                        XLSUtil.copyRow(sourceSheet, targetSheet, parentRowNum, newRowNum, parentStyles.get(parentOutlineLevel));
                        newRowNum++;
                    }
                }
                // Copy main row suitable for filter
                XLSUtil.copyRow(sourceSheet, targetSheet, i, newRowNum, rowStyles.get(outlineLevel));
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
}
