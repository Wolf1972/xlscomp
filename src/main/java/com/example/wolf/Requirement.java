package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Requirement {

    static final int HEADER_LAST_ROW = 1; // Header row index (starts from 0)
    static final int LAST_COMMON_COLUMN = 12; // Last copying column index (starts from 0) - to prevent copying service columns

    String id; // Requirement id
    private Integer row; // Excel sheet row num (when requirement loads from Excel sheet)
    private Integer[] parentRows; // List of all parents Excel rows

    private Integer level; // Requirement level
    private String name; // Requirement
    private String priority; // Requirement priority
    private String done; // Requirement has realised
    private String reference; // Requirement from other source (MarxxWeb)
    private String integration; // Integration requirement
    private String comment; // Comment for requirement
    private String linked; // Linked requirement
    private String version; // Plan to realised in version (13)
    private String release; // Plan to realized in release (14)
    private String questions; // Work questions for requirement (15)
    private String source_req; // Requirement in source (16)
    private String customize; // Requirement realizes by customize (17)
    private String tt; // Team track task (18)
    private String trello; // Trello task (19)

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;

        Requirement that = (Requirement) o;

        if (!id.equals(that.id)) return false;
        if (row != null ? !row.equals(that.row) : that.row != null) return false;
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
        if (questions != null ? !questions.equals(that.questions) : that.questions != null) return false;
        if (source_req != null ? !source_req.equals(that.source_req) : that.source_req != null) return false;
        if (customize != null ? !customize.equals(that.customize) : that.customize != null) return false;
        if (tt != null ? !tt.equals(that.tt) : that.tt != null) return false;
        return trello != null ? trello.equals(that.trello) : that.trello == null;
    }

    @Override
    public int hashCode() {
        int result = id.hashCode();
        result = 31 * result + (row != null ? row.hashCode() : 0);
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
        result = 31 * result + (source_req != null ? source_req.hashCode() : 0);
        result = 31 * result + (customize != null ? customize.hashCode() : 0);
        result = 31 * result + (tt != null ? tt.hashCode() : 0);
        result = 31 * result + (trello != null ? trello.hashCode() : 0);
        return result;
    }

    /**
     * Fills object fields from XLSX row
     * @param xrow - Excel XLSX row
     */
    void loadFromRow(XSSFRow xrow) {
        int cells = xrow.getLastCellNum();
        row = xrow.getRowNum();

        if (cells > 0) level = safeLoadInteger(xrow, 0); // Requirement level
        if (cells > 1) name = safeLoadString(xrow, 1); // Requirement
        if (name.contains("\\")) {
            System.out.println("WARNING. Name for row " + (row + 1) + " contains \\");
            name = name.replace("\\", "");
        }
        if (cells > 2) priority = safeLoadString(xrow, 2); // Requirement priority
        if (cells > 3) done = safeLoadString(xrow,3); // Requirement has realised
        if (cells > 4) reference = safeLoadString(xrow, 4); // Requirement from other source (MarxxWeb)

        if (cells > 6) integration = safeLoadString(xrow, 6); // Integration requirement
        if (cells > 7) comment = safeLoadString(xrow, 7); // Comment for requirement
        if (cells > 8) linked = safeLoadString(xrow,8); // Linked requirement

        if (cells > 13) version = safeLoadString(xrow,13); // Plan to realised in version
        if (cells > 14) release = safeLoadString(xrow, 14); // Plan to realized in release
        if (cells > 15) questions = safeLoadString(xrow,15); // Work questions for requirement
        if (cells > 16) source_req = safeLoadString(xrow,16); // Requirement in source
        if (cells > 17) customize = safeLoadString(xrow,17); // Requirement realizes by customize
        if (cells > 18) tt = safeLoadString(xrow,18); // Team track task
        if (cells > 19) trello = safeLoadString(xrow,19); // Trello task

        // Id evaluation - get all parent nodes
        int outlineLevel = xrow.getOutlineLevel();
        if ((outlineLevel + 1) != level) {
            System.out.println("ERROR. Value specified in the first column for row " + (row + 1) + " has level " + level + " mismatches with outline level: " + (outlineLevel + 1));
        }
        int parentRow = row;
        parentRows = new Integer[outlineLevel]; // Array for all parent rows
        StringBuilder fullPath = new StringBuilder("|" + name);
        for (int i = outlineLevel - 1; i >= 0 ; i--) {
            while (true) {
                parentRow--;
                if (parentRow < 0) {
                    System.out.println("ERROR. Can't find full path for row " + (row + 1));
                    break;
                }
                XSSFRow prevRow = xrow.getSheet().getRow(parentRow);
                int prevOutlineLevel = prevRow.getOutlineLevel();
                if (prevOutlineLevel == i) {
                    if (i < outlineLevel - 1) fullPath.insert(0, "\\");
                    fullPath.insert(0, prevRow.getCell(1).getStringCellValue());
                    parentRows[prevOutlineLevel] = parentRow;
                    break;
                }
                else if (prevOutlineLevel < i) {
                    System.out.println("ERROR. Outline levels sequence violation for row " + (row + 1));
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
    private String safeLoadString(XSSFRow row, int column) {
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
    private Integer safeLoadInteger(XSSFRow row, int column) {
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
                    // Do nothing;
                } else if (type == CellType.NUMERIC) {
                    result = new Integer((int) Math.round(cell.getNumericCellValue()));
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
        cell = row.createCell(16); cell.setCellValue(source_req); // Source requirement
        cell = row.createCell(17); cell.setCellValue(customize); // Realizes by customize
        cell = row.createCell(18); cell.setCellValue(tt); // Team track task
        cell = row.createCell(19); cell.setCellValue(trello); // Trello task
    }

}
