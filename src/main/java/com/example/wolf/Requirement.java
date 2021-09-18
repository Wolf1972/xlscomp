package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Requirement {

    String id; // Requirement id
    Long level; // Requirement level
    String name; // Requirement
    String priority; // Requirement priority
    String done; // Requirement has realised
    String reference; // Requirement from other source (MarxxWeb)
    String integration; // Integration requirement
    String comment; // Comment for requirement
    String linked; // Linked requirement
    String version; // Plan to realised in version
    String release; // Plan to realized in release
    String questions; // Work questions for requirement

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Requirement that = (Requirement) o;

        if (!id.equals(that.id)) return false;
        if (level != null ? !level.equals(that.level) : that.level != null) return false;
        if (name != null ? !name.equals(that.name) : that.name != null) return false;
        if (priority != null ? !priority.equals(that.priority) : that.priority != null) return false;
        if (done != null ? !done.equals(that.done) : that.done != null) return false;
        if (reference != null ? !reference.equals(that.reference) : that.reference != null) return false;
        if (integration != null ? !integration.equals(that.integration) : that.integration != null) return false;
        if (comment != null ? !comment.equals(that.comment) : that.comment != null) return false;
        if (linked != null ? !linked.equals(that.linked) : that.linked != null) return false;
        if (version != null ? !version.equals(that.version) : that.version != null) return false;
        if (release != null ? !release.equals(that.release) : that.release != null) return false;
        return questions != null ? questions.equals(that.questions) : that.questions == null;
    }

    @Override
    public int hashCode() {
        int result = id.hashCode();
        result = 31 * result + (level != null ? level.hashCode() : 0);
        result = 31 * result + (name != null ? name.hashCode() : 0);
        result = 31 * result + (priority != null ? priority.hashCode() : 0);
        result = 31 * result + (done != null ? done.hashCode() : 0);
        result = 31 * result + (reference != null ? reference.hashCode() : 0);
        result = 31 * result + (integration != null ? integration.hashCode() : 0);
        result = 31 * result + (comment != null ? comment.hashCode() : 0);
        result = 31 * result + (linked != null ? linked.hashCode() : 0);
        result = 31 * result + (version != null ? version.hashCode() : 0);
        result = 31 * result + (release != null ? release.hashCode() : 0);
        result = 31 * result + (questions != null ? questions.hashCode() : 0);
        return result;
    }

    /**
     * Fills object fields from XLSX row
     * @param row - Excel XLSX row
     */
    public void loadFromRow(XSSFRow row) {
        int cells = row.getLastCellNum();

        if (cells > 0) level = safeLoadInteger(row, 0); // Requirement level
        if (cells > 1) name = safeLoadString(row, 1); // Requirement
        if (name.contains("\\")) {
            System.out.println("WARNING. Name for row " + (row.getRowNum() + 1) + " contains \\");
            name = name.replace("\\", "");
        }
        if (cells > 2) priority = safeLoadString(row, 2); // Requirement priority
        if (cells > 3) done = safeLoadString(row,3); // Requirement has realised
        if (cells > 4) reference = safeLoadString(row, 4); // Requirement from other source (MarxxWeb)

        if (cells > 6) integration = safeLoadString(row, 6); // Integration requirement
        if (cells > 7) comment = safeLoadString(row, 7); // Comment for requirement
        if (cells > 8) linked = safeLoadString(row,8); // Linked requirement

        if (cells > 13) version = safeLoadString(row,13); // Plan to realised in version
        if (cells > 14) release = safeLoadString(row, 14); // Plan to realized in release
        if (cells > 15) questions = safeLoadString(row,15); // Work questions for requirement

        // Id evaluation - get all parent nodes
        int outlineLevel = row.getOutlineLevel();
        if ((outlineLevel + 1) != level) {
            System.out.println("ERROR. Row " + (row.getRowNum() + 1) + " has level " + level + " mismatches with outline: " + (outlineLevel + 1));
        }
        int rowNum = row.getRowNum();
        StringBuilder fullPath = new StringBuilder("|" + name);
        for (int i = outlineLevel - 1; i >= 0 ; i--) {
            while (true) {
                rowNum--;
                if (rowNum < 0) {
                    System.out.println("ERROR. Can't find full path for row " + (row.getRowNum() + 1));
                    break;
                }
                XSSFRow prevRow = row.getSheet().getRow(rowNum);
                int prevOutlineLevel = prevRow.getOutlineLevel();
                if (prevOutlineLevel == i) {
                    if (i < outlineLevel - 1) fullPath.insert(0, "\\");
                    fullPath.insert(0, prevRow.getCell(1).getStringCellValue());
                    break;
                }
                else if (prevOutlineLevel < i) {
                    System.out.println("ERROR. Outline levels sequence violation for row " + (row.getRowNum() + 1));
                    break;
                }
            }
        }
        fullPath.insert(0, "\\");
        id = fullPath.toString();
    }

    /**
     * Returns string from cell (even if column contains number or error)
     * @param row - XLSX row
     * @param column - column number (starts from 0)
     * @return - string
     */
    public String safeLoadString(XSSFRow row, int column) {
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
                    result = cell.getCellFormula();
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
    public Long safeLoadInteger(XSSFRow row, int column) {
        Long result = 0L;

        if (row != null) {
            try {
                XSSFCell cell = row.getCell(column);
                CellType type = cell.getCellType();
                if (type == CellType.BLANK) {
                    // Do nothing
                } else if (type == CellType.BOOLEAN) {
                    Boolean logical = cell.getBooleanCellValue();
                    result = logical ? 1L : 0L;
                } else if (type == CellType.ERROR) {
                    // Do nothing;
                } else if (type == CellType.FORMULA) {
                    // Do nothing;
                } else if (type == CellType.NUMERIC) {
                    result = (long) Math.round(cell.getNumericCellValue());
                } else if (type == CellType.STRING) {
                    result = Long.parseLong(cell.getStringCellValue());
                }
            }
            catch (Exception e) {
                System.out.println("ERROR while loading row " + (row.getRowNum() + 1) + ", column " + (column + 1) + ": " + e.getMessage());
            }
        }
        return result;
    }

    /**
     * Fills XLSX row from object
     * @param row - XLSX row
     */
    public void saveToRow(XSSFRow row) {
        XSSFCell cell = row.createCell(0); cell.setCellValue(level); // Requirement level
        cell = row.createCell(1); cell.setCellValue(name); // Requirement
        cell = row.createCell(2); cell.setCellValue(priority); // Requirement priority
        cell = row.createCell(3); cell.setCellValue(done); // Requirement has realised
        cell = row.createCell(4); cell.setCellValue(reference); // Requirement from other source (MarxxWeb)

        cell = row.createCell(6); cell.setCellValue(integration); // Integration requirement
        cell = row.createCell(7); cell.setCellValue(comment); // Comment for requirement
        cell = row.createCell(8); cell.setCellValue(linked); // Linked requirement

        cell = row.createCell(13); cell.setCellValue(version); // Plan to realised in version
        cell = row.createCell(14); cell.setCellValue(release); // Plan to realized in release
        cell = row.createCell(15); cell.setCellValue(questions); // Work questions for requirement
    }

}
