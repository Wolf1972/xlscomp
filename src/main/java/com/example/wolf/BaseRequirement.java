package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class BaseRequirement {
    /**
     * Returns string from cell (even if column contains number or error)
     * @param row - XLSX row
     * @param column - column number (starts from 0)
     * @return - string
     */
    String safeLoadString(XSSFRow row, int column) {
        String result = "";

        if (row != null) {
            try {
                XSSFCell cell = row.getCell(column);
                CellType type = cell.getCellType();
                if (type == CellType.BLANK) {
                    // Do nothing
                } else if (type == CellType.BOOLEAN) {
                    Boolean logical = cell.getBooleanCellValue();
                    result = logical.toString();
                } else if (type == CellType.ERROR) {
                    result = cell.getErrorCellString();
                } else if (type == CellType.FORMULA) {
                    try {
                        result = cell.getStringCellValue();
                    }
                    catch (Exception e) { // Formula can returns numeric value
                        Double number = cell.getNumericCellValue();
                        result = number.toString();
                    }
                } else if (type == CellType.NUMERIC) {
                    Double number = cell.getNumericCellValue();
                    result = number.toString();
                } else if (type == CellType.STRING) {
                    result = cell.getStringCellValue();
                }
            }
            catch (Exception e) {
                System.out.println("ERROR while loading row " + (row.getRowNum() + 1) + ", column " + column + ": " + e.getMessage());
            }
        }
        return result;
    }

    /**
     * Returns integer from cell (even if column contains string or error)
     * @param row - XLSX row
     * @param column - column number (starts from 0)
     * @return - long value
     */
    Integer safeLoadInteger(XSSFRow row, int column) {
        Integer result = 0;

        if (row != null) {
            try {
                XSSFCell cell = row.getCell(column);
                CellType type = cell.getCellType();
                if (type == CellType.BLANK) {
                    // Do nothing
                } else if (type == CellType.BOOLEAN) {
                    Boolean logical = cell.getBooleanCellValue();
                    result = logical ? 1 : 0;
                } else if (type == CellType.ERROR) {
                    // Do nothing;
                } else if (type == CellType.FORMULA) {
                    try {
                        result = (int) Math.round(cell.getNumericCellValue());
                    }
                    catch (Exception e) { // Formula can returns string value
                        result = Integer.parseInt(cell.getStringCellValue());
                    }
                } else if (type == CellType.NUMERIC) {
                    result = (int) Math.round(cell.getNumericCellValue());
                } else if (type == CellType.STRING) {
                    result = Integer.parseInt(cell.getStringCellValue());
                }
            }
            catch (Exception e) {
                System.out.println("ERROR while loading row " + (row.getRowNum() + 1) + ", column " + (column + 1) + ": " + e.getMessage());
            }
        }
        return result;
    }

    /**
     * Returns integer from cell (even if column contains string or error)
     * @param row - XLSX row
     * @param column - column number (starts from 0)
     * @return - long value
     */
    Double safeLoadDouble(XSSFRow row, int column) {
        Double result = 0.0;

        if (row != null) {
            try {
                XSSFCell cell = row.getCell(column);
                CellType type = cell.getCellType();
                if (type == CellType.BLANK) {
                    // Do nothing
                } else if (type == CellType.BOOLEAN) {
                    Boolean logical = cell.getBooleanCellValue();
                    result = logical ? 1.0 : 0.0;
                } else if (type == CellType.ERROR) {
                    // Do nothing;
                } else if (type == CellType.FORMULA) {
                    try {
                        result = cell.getNumericCellValue();
                    }
                    catch (Exception e) { // Formula can returns string value
                        result = Double.parseDouble(cell.getStringCellValue());
                    }
                } else if (type == CellType.NUMERIC) {
                    result = cell.getNumericCellValue();
                } else if (type == CellType.STRING) {
                    result = Double.parseDouble(cell.getStringCellValue());
                }
            }
            catch (Exception e) {
                System.out.println("ERROR while loading row " + (row.getRowNum() + 1) + ", column " + (column + 1) + ": " + e.getMessage());
            }
        }
        return result;
    }
}
