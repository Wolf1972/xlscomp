package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

public class XLSUtil {
    /**
     * Copying first sheet from specified source file into specified target sheet
     * Defines outline level while copying for grouping rows
     * Collects styles while copying: several header rows separately with common styles for requirement rows
     * Grouping rows with outline levels after copying
     * @param sourceFile - source file name
     * @param targetSheet - target sheet
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     */
    static void copySheet(String sourceFile, XSSFSheet targetSheet, int maxColumn) {

        class Group { // Item for rows grouping
            private int start;
            private int end;
            private int level;
            private boolean closed;
            private Group(int start, int end, int level) {
                this.start = start;
                this.end = end;
                this.level = level;
                this.closed = false;
            }
        }

        ArrayList<Group> groups = new ArrayList<>();

        // Common styles for all group levels (with outline levels)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        try {

            XSSFWorkbook sourceBook = new XSSFWorkbook(new FileInputStream(sourceFile));
            XSSFSheet sourceSheet = sourceBook.getSheetAt(0);
            int oldOutlineLevel = 0;
            int lastRow = sourceSheet.getLastRowNum();

            for (int i = 0; i <= lastRow; i++) {

                int outlineLevel = sourceSheet.getRow(i).getOutlineLevel();

                if (i > Requirement.HEADER_LAST_ROW && outlineLevel >= 0) {

                    int specifiedOutlineLevel = (int) sourceSheet.getRow(i).getCell(0).getNumericCellValue();
                    if (outlineLevel + 1 != specifiedOutlineLevel) {
                        System.out.println("ERROR in row " + (i + 1) + ". Real row outline level " + (outlineLevel + 1) + " doesn't suite with level has specified in first column: " + specifiedOutlineLevel);
                    }

                    if (oldOutlineLevel != outlineLevel) {
                        if (oldOutlineLevel < outlineLevel) { // Dive! Dive! Dive
                            groups.add(new Group(i, 0, outlineLevel));
                        } else { // Surfacing!
                            for (int g = outlineLevel; g <= oldOutlineLevel; g++) { // May be close several groups in the same time
                                for (Group group : groups) {
                                    if (!group.closed && group.level == g + 1) {
                                        group.end = i - 1;
                                        group.closed = true;
                                    }
                                }
                            }
                        }
                        oldOutlineLevel = outlineLevel;
                    }
                }

                ArrayList<XSSFCellStyle> styles = new ArrayList<>(); // Styles for current row

                if (i <= Requirement.HEADER_LAST_ROW || !groupStyles.containsKey(outlineLevel)) {
                    // Styles for header row or row with unknown outline level
                    // Copy style from old cell and apply to new cell: all styles after specified row are common - takes it from array
                    for (int j = 0; j < sourceSheet.getRow(i).getLastCellNum(); j++) {
                        XSSFCell cell = sourceSheet.getRow(i).getCell(j);
                        if (cell != null) { // Sometimes we can not obtain some cell even if j < getLastCellNum (possible - merged cells?)
                            XSSFCellStyle newCellStyle = targetSheet.getWorkbook().createCellStyle();
                            newCellStyle.cloneStyleFrom(cell.getCellStyle());
                            styles.add(newCellStyle);
                        }
                    }
                    if (i > Requirement.HEADER_LAST_ROW) { // For regular rows with requirement add common style for outline level
                        groupStyles.put(outlineLevel, styles);
                    }
                }
                else {
                    styles = groupStyles.get(outlineLevel); // Use common style has already defined
                }
                copyRow(sourceSheet, targetSheet, i, i, maxColumn, styles);
            }

            sourceBook.close();

            for (Group group : groups) { // Close all unclosed groups with last row
                if (!group.closed) {
                    group.end = lastRow;
                    group.closed = true;
                }
            }

        }
        catch (IOException e) {
            System.out.println("Error while reading source sheet: " + sourceFile);
        }

        targetSheet.setRowSumsBelow(false); // Set group header at the top of group
        for (Group group : groups) {
            // System.out.println(group.start + " : " + group.end + " - " + group.level);
            targetSheet.groupRow(group.start, group.end);
        }

    }

    /**
     * Copying one specified row (by position) from specified source sheet to specified target sheet
     * If row with specified number already exists in target row, then row inserts with scroll rows below
     * Sets specified styles for target row
     * Based on some stackoverflow topics
     * @param sourceWorksheet - source sheet
     * @param targetWorksheet = target sheet
     * @param sourceRowNum = source row number
     * @param targetRowNum - destination row number
     * @param maxColumn - last column to copy (to prevent copying service secured columns), when = 0 - copying all columns from row
     * @param columnStyles - styles for all columns
     */
    static void copyRow(XSSFSheet sourceWorksheet, XSSFSheet targetWorksheet,
                        int sourceRowNum, int targetRowNum,
                        int maxColumn,
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

            // Do not copy service columns
            if ((maxColumn > 0) && (i > maxColumn)) {
                continue;
            }
            // If the old cell is null jump to next cell (may be merged cells?)
            if (oldCell == null) {
                continue;
            }

            if (columnStyles != null && i < columnStyles.size()) {
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
