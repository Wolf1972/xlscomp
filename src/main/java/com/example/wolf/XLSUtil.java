package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class XLSUtil {
    /**
     * Copying first sheet from specified source file into specified target sheet
     * Defines outline level while copying for grouping rows
     * Collects styles while copying: several header rows separately with common styles for requirement rows
     * Grouping rows with outline levels after copying
     * @param sourceFile - source file name
     * @param targetSheet - target sheet
     * @param columnsQty - quantity of columns to copy (to prevent output service columns), when 0 - all columns copy
     */
    static void copySheet(String sourceFile, XSSFSheet targetSheet, int columnsQty) {

        // Common styles for all group levels (with outline levels)
        HashMap<Integer, ArrayList<XSSFCellStyle>> groupStyles = new HashMap<>();

        try {

            XSSFWorkbook sourceBook = new XSSFWorkbook(new FileInputStream(sourceFile));
            XSSFSheet sourceSheet = sourceBook.getSheetAt(0);

            copyHeader(sourceSheet, targetSheet, columnsQty); // Copy header

            int lastRow = sourceSheet.getLastRowNum();

            ArrayList<XSSFCellStyle> styles = new ArrayList<>();

            for (int i = Requirement.HEADER_LAST_ROW + 1; i <= lastRow; i++) {

                int outlineLevel = sourceSheet.getRow(i).getOutlineLevel();
                if (!groupStyles.containsKey(outlineLevel)) {
                    styles = getRowStyles(sourceSheet, i, columnsQty);
                    styles = cloneRowStyles(targetSheet, styles);
                    groupStyles.put(outlineLevel, styles);
                }
                else {
                    styles = groupStyles.get(outlineLevel); // Use common style has already defined
                }
                copyRow(sourceSheet, targetSheet, true, i, i, columnsQty, null, false, styles, null, null);
            }

            sourceBook.close();

        }
        catch (IOException e) {
            System.out.println("Error while reading source sheet: " + sourceFile);
        }
    }

    /**
     * Method copies only header of sheet
     * @param sourceSheet - source sheet
     * @param targetSheet - target sheet
     * @param columnsQty - quantity of columns to copy (to prevent output service columns), when 0 - all columns copy
     */
    static void copyHeader(XSSFSheet sourceSheet, XSSFSheet targetSheet, int columnsQty) {

        for (int i = 0; i <= Requirement.HEADER_LAST_ROW; i++) {
            ArrayList<XSSFCellStyle> styles = getRowStyles(sourceSheet, i, columnsQty);
            styles = cloneRowStyles(targetSheet, styles);
            copyRow(sourceSheet, targetSheet, true, i, i, columnsQty, null, false, styles, null, null);
        }
    }

    /**
     * Returns array with all column styles
     * @param sourceSheet - sheet to get styles
     * @param rowNum - row to get styles
     * @param columnsQty - quantity of columns to copy (to prevent output service columns), when 0 - all columns copy
     * @return - array with styles
     */
    static ArrayList<XSSFCellStyle> getRowStyles(XSSFSheet sourceSheet, int rowNum , int columnsQty) {
        ArrayList<XSSFCellStyle> styles = new ArrayList<>(); // Styles for current row
        XSSFRow row = sourceSheet.getRow(rowNum);
        if (row != null) {
            for (int j = 0; j < row.getLastCellNum(); j++) {
                if (columnsQty > 0 && j > columnsQty) break;
                XSSFCell cell = row.getCell(j);
                if (cell != null) { // Sometimes we can not obtain some cell even if j < getLastCellNum (possible - merged cells?)
                    styles.add(cell.getCellStyle());
                }
            }
        }
        else {
            System.out.println("ERROR. There is no source row " + rowNum + " in source sheet" + sourceSheet.getSheetName());
        }
        return styles;
    }

    /**
     * Clones styles array into target sheet
     * @param sheet - target sheet
     * @param styles - styles
     * @return - array with styles were cloned
     */
    static ArrayList<XSSFCellStyle> cloneRowStyles(XSSFSheet sheet, ArrayList<XSSFCellStyle> styles) {
        ArrayList<XSSFCellStyle> newStyles = new ArrayList<>();
        if (styles != null) {
            for (int i = 0; i < styles.size(); i++) {
                XSSFCellStyle newCellStyle = sheet.getWorkbook().createCellStyle();
                newCellStyle.cloneStyleFrom(styles.get(i));
                newStyles.add(newCellStyle);
            }
        }
        return newStyles;
    }

    /**
     * Modifies cell styles according with mode (deleted, addede, changed or parent mode)
     * @param sheet - target sheet
     * @param row - row for styles modify
     * @param mark - row mark type (DELETED, CHANGED, ADDED, PARENT or null)
     */
    static ArrayList<XSSFCellStyle> modifyCellStyles(XSSFSheet sheet, XSSFRow row, MarkRowType mark) {
        ArrayList<XSSFCellStyle> styles = new ArrayList<>();
        for (int j = 0; j < row.getLastCellNum(); j++) {
            XSSFCell cell = row.getCell(j);
            XSSFCellStyle newCellStyle = sheet.getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(cell.getCellStyle());
            Font font = sheet.getWorkbook().createFont();
            if (mark != null) {
                switch (mark) {
                    case DELETED: {
                        font.setBold(true);
                        font.setStrikeout(true);
                        break;
                    }
                    case ADDED: {
                        font.setBold(true);
                        font.setColor(IndexedColors.RED.getIndex());
                        break;
                    }
                    case CHANGED: {
                        font.setBold(true);
                        font.setColor(IndexedColors.BLUE.getIndex());
                        break;
                    }
                    case PARENT: {
                        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
                        break;
                    }
                }
                newCellStyle.setFont(font);
                styles.add(newCellStyle);
            }
        }
        return styles;
    }

    /**
     * Function check for according between row outline level and group number (specified in 1st column)
     * List of errors outputs to console
     * @param sheet - sheet for processing
     * @param startRow - start row for check (till end row of sheet)
     * @return - when true - all Ok, when false - sheet contain errors
     */
    static boolean checkSheetOutline(XSSFSheet sheet, int startRow) {
        boolean result = true;
        int lastRow = sheet.getLastRowNum();

        for (int i = startRow; i <= lastRow; i++) {

            int outlineLevel = (int) sheet.getRow(i).getOutlineLevel();
            int specifiedOutlineLevel = (int) sheet.getRow(i).getCell(0).getNumericCellValue();

            if ((outlineLevel + 1) != specifiedOutlineLevel) {
                System.out.println("ERROR in row " + (i + 1) + ". Real row outline level " + (outlineLevel + 1) + " doesn't suite with level has specified in first column: " + specifiedOutlineLevel);
                result = false;
            }
        }
        return result;
    }

    /**
     * Function groups all rows in specified sheet by outline level (outline level takes from 1st column)
     * @param sheet - sheet to group by
     * @param startRow - start row for grouping (till end row of sheet)
     */
    static void groupSheet(XSSFSheet sheet, int startRow) {

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

        int oldOutlineLevel = 0;
        int lastRow = sheet.getLastRowNum();

        for (int i = startRow; i <= lastRow; i++) {

            int outlineLevel = (int) sheet.getRow(i).getCell(0).getNumericCellValue() - 1;

            if (outlineLevel >= 0) {

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
        }
        for (Group group : groups) { // Close all unclosed groups with last row
            if (!group.closed) {
                group.end = lastRow;
                group.closed = true;
            }
        }

        sheet.setRowSumsBelow(false); // Set group header at the top of group
        for (Group group : groups) {
            // System.out.println(group.start + " : " + group.end + " - " + group.level);
            sheet.groupRow(group.start, group.end);
        }
    }

    /**
     * Copying one specified row (by position) from specified source sheet to specified target sheet
     * If row with specified number already exists in target row, then row inserts with scroll rows below
     * Sets specified styles for target row
     * Based on some stackoverflow topics
     * @param sourceWorksheet - source sheet
     * @param targetWorksheet - target sheet
     * @param isAppend - append mode
     * @param sourceRowNum = source row number
     * @param targetRowNum - destination row number
     * @param columnsQty - quantity of columns to copy (to prevent output service columns), when 0 - all columns copy
     * @param onlyColumns - copy only rows specified in this array (when null - this filter is not applying); column indexes for source sheet
     * @param isOnlyEmpty - when true: copy only empty cells in target sheet, when false - any cells
     * @param columnStyles - styles for all columns (may be null); styles for columns for source sheet
     * @param sourceDescriber - source sheet column describer (if only structure difference between source and target)
     * @param targetDescriber - target sheet column describer (if only structure difference between source and target)
     * @param isDebug - debug mode
     */
    static void copyRow(XSSFSheet sourceWorksheet, XSSFSheet targetWorksheet,
                        boolean isAppend,
                        int sourceRowNum, int targetRowNum,
                        int columnsQty,
                        List<Integer> onlyColumns,
                        boolean isOnlyEmpty,
                        ArrayList<XSSFCellStyle> columnStyles,
                        RequirementColumnDescriber sourceDescriber,
                        RequirementColumnDescriber targetDescriber,
                        boolean isDebug) {

        // Get the source / new row
        XSSFRow newRow = targetWorksheet.getRow(targetRowNum);
        XSSFRow sourceRow = sourceWorksheet.getRow(sourceRowNum);

        StringBuilder debugStr = new StringBuilder("Merge row ");
        debugStr.append(sourceRowNum + " " + sourceRow.getCell(1).getStringCellValue() + ": ");

        if (isAppend) {
            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null) {
                targetWorksheet.shiftRows(targetRowNum, targetWorksheet.getLastRowNum(), 1);
            } else {
                newRow = targetWorksheet.createRow(targetRowNum);
            }
        }
        else
            newRow = targetWorksheet.getRow(targetRowNum);

        // Loop through source columns to add to new row
        for (int oldIdx = 0; oldIdx < sourceRow.getLastCellNum(); oldIdx++) {

            // Columns mapping
            int newIdx = oldIdx;
            if (sourceDescriber != null && targetDescriber != null) {
                RequirementFieldType oldColumnType = sourceDescriber.getField(oldIdx);
                Integer newColumnIdx = targetDescriber.getColumn(oldColumnType); // Columns mapping
                if (newColumnIdx != null)
                    newIdx = newColumnIdx;
                else
                    continue;
            }

            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(oldIdx);

            XSSFCell newCell = null;
            if (newRow.getLastCellNum() >= newIdx) newCell = newRow.getCell(newIdx);

            if (oldCell == null) continue; // If the old cell is null, then jump to next cell (may be merged cells?)
            if (newCell == null) newCell = newRow.createCell(newIdx);

            // Do not copy service columns
            if ((columnsQty > 0) && (newIdx > columnsQty)) {
                continue;
            }

            if (onlyColumns != null) {
                if (!onlyColumns.contains(oldIdx)) continue;
            }

            if (columnStyles != null && oldIdx < columnStyles.size()) {
                XSSFCellStyle newCellStyle = columnStyles.get(oldIdx);
                newCell.setCellStyle(newCellStyle);
            }

            targetWorksheet.setColumnWidth(newIdx, sourceWorksheet.getColumnWidth(oldIdx));

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            boolean isNeedCopy = true;
            if (isOnlyEmpty) { // When we are copying only empty target cells
                isNeedCopy = false;
                CellType newType = newCell.getCellType();
                if (newType == CellType.BLANK) isNeedCopy = true;
                if (!isNeedCopy && (newType == CellType.STRING) && newCell.getRichStringCellValue().getString().isEmpty()) isNeedCopy = true;
                if (!isNeedCopy && (newType == CellType.NUMERIC) && (newCell.getNumericCellValue() == 0)) isNeedCopy = true;
            }
            if (!isNeedCopy) continue;

            debugStr.append("[" + oldIdx + " -> " + newIdx + "]: ");

            // Set the cell data value
            CellType oldType = oldCell.getCellType();
            if (oldType == CellType.BLANK) {
                newCell.setCellType(oldType);
                String old = oldCell.getStringCellValue();
                newCell.setCellValue(old);
                debugStr.append(old);
            }
            else if (oldType == CellType.BOOLEAN) {
                newCell.setCellType(oldType);
                boolean old = oldCell.getBooleanCellValue();
                newCell.setCellValue(old);
                debugStr.append(old);
            }
            else if (oldType == CellType.ERROR) {
                newCell.setCellType(oldType);
                byte old = oldCell.getErrorCellValue();
                newCell.setCellErrorValue(old);
                debugStr.append(old);
            }
            else if (oldType == CellType.FORMULA) {
                String old = oldCell.getCellFormula();
                newCell.setCellFormula(old);
                debugStr.append(old);
            }
            else if (oldType == CellType.NUMERIC) {
                newCell.setCellType(oldType);
                double old = oldCell.getNumericCellValue();
                newCell.setCellValue(old);
                debugStr.append(old);
            }
            else if (oldType == CellType.STRING) {
                newCell.setCellType(oldType);
                // XSSFRichTextString old = oldCell.getRichStringCellValue();
                String old = oldCell.getStringCellValue();
                newCell.setCellValue(old);
                debugStr.append(old);
            }

            debugStr.append(" ");

        }

        if (isDebug) System.out.println(debugStr.toString());

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

    /**
     * Previous method by default - without debug
     */
    static void copyRow(XSSFSheet sourceWorksheet, XSSFSheet targetWorksheet,
                        boolean isAppend,
                        int sourceRowNum, int targetRowNum,
                        int columnsQty,
                        List<Integer> onlyColumns,
                        boolean isOnlyEmpty,
                        ArrayList<XSSFCellStyle> columnStyles,
                        RequirementColumnDescriber sourceDescriber,
                        RequirementColumnDescriber targetDescriber) {
        copyRow(sourceWorksheet, targetWorksheet, isAppend, sourceRowNum, targetRowNum, columnsQty,
                onlyColumns, isOnlyEmpty, columnStyles, sourceDescriber, targetDescriber, false);
    };
}
