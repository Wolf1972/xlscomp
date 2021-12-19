package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

public class Requirement extends BaseRequirement {

    static final int HEADER_LAST_ROW = 1; // Header row index (starts from 0)

    static final int MXWEB_RELEASE = 17; // MxWeb release column (column R - index 17 from 0)

    String id; // Requirement id
    private Integer row; // Excel sheet row num (when requirement loads from Excel sheet)
    private Integer[] parentRows; // List of all parents Excel rows
    // Public columns
    private Integer level; // A(0): Requirement level
    private String name; // B(1): Requirement
    private String priority; // C(2): Requirement priority
    private String done; // D(3): Requirement has realised
    private String reference; // E(4): Requirement from other source (MarxWeb)
    private String new_req; // F(5): New requirement flag
    private String integration; // G(6): Integration requirement
    private String service; // H(7): Integration service requirement
    private String comment; // I(8): Comment for requirement
    private String linked; // J(9): Linked requirement
    private String curr_status; // K(10): Current status
    private String type; // L(11): Requirement type
    private String source; // M(12): Requirement source
    private String foundation; // N(13): Requirement foundation
    // Private columns
    private String version; // O(14): Plan to realised in version
    private String release; // P(15): Plan to realized in release
    private String questions; // Q(16): Work questions for requirement
    private String source_req; // R(17): Requirement in source
    private String tt; // S(18): Team track task
    private String trello; // T(19): Trello task
    private String primary; // U(20): Primary responsible
    private String secondary; // V(21): Secondary responsible
    private String risk; // W(22): Risk
    private String risk_desc; // X(23): Risk description

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
        if (new_req != null ? !new_req.equals(that.new_req) : that.new_req != null) return false;
        if (integration != null ? !integration.equals(that.integration) : that.integration != null) return false;
        if (service != null ? !service.equals(that.service) : that.service != null) return false;
        if (comment != null ? !comment.equals(that.comment) : that.comment != null) return false;
        if (linked != null ? !linked.equals(that.linked) : that.linked != null) return false;
        if (curr_status != null ? !curr_status.equals(that.curr_status) : that.curr_status != null) return false;
        if (type != null ? !type.equals(that.type) : that.type != null) return false;
        if (source != null ? !source.equals(that.source) : that.source != null) return false;
        if (foundation != null ? !foundation.equals(that.foundation) : that.foundation != null) return false;

        if (version != null ? !version.equals(that.version) : that.version != null) return false;
        if (release != null ? !release.equals(that.release) : that.release != null) return false;
        if (questions != null ? !questions.equals(that.questions) : that.questions != null) return false;
        if (source_req != null ? !source_req.equals(that.source_req) : that.source_req != null) return false;
        if (tt != null ? !tt.equals(that.tt) : that.tt != null) return false;
        if (trello != null ? !trello.equals(that.trello) : that.trello != null) return false;
        if (primary != null ? !primary.equals(that.primary) : that.primary != null) return false;
        if (secondary != null ? !secondary.equals(that.secondary) : that.secondary != null) return false;
        if (risk != null ? !risk.equals(that.risk) : that.risk != null) return false;
        return risk_desc != null ? risk_desc.equals(that.risk_desc) : that.risk_desc == null;

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
        result = 31 * result + (new_req != null ? new_req.hashCode() : 0);
        result = 31 * result + (integration != null ? integration.hashCode() : 0);
        result = 31 * result + (service != null ? service.hashCode() : 0);
        result = 31 * result + (comment != null ? comment.hashCode() : 0);
        result = 31 * result + (linked != null ? linked.hashCode() : 0);
        result = 31 * result + (curr_status != null ? curr_status.hashCode() : 0);
        result = 31 * result + (type != null ? type.hashCode() : 0);
        result = 31 * result + (source != null ? source.hashCode() : 0);
        result = 31 * result + (foundation != null ? foundation.hashCode() : 0);

        result = 31 * result + (version != null ? version.hashCode() : 0);
        result = 31 * result + (release != null ? release.hashCode() : 0);
        result = 31 * result + (questions != null ? questions.hashCode() : 0);
        result = 31 * result + (source_req != null ? source_req.hashCode() : 0);
        result = 31 * result + (tt != null ? tt.hashCode() : 0);
        result = 31 * result + (trello != null ? trello.hashCode() : 0);
        result = 31 * result + (primary != null ? primary.hashCode() : 0);
        result = 31 * result + (secondary != null ? secondary.hashCode() : 0);
        result = 31 * result + (risk != null ? risk.hashCode() : 0);
        result = 31 * result + (risk_desc != null ? risk_desc.hashCode() : 0);

        return result;
    }

    public Integer getRow() {
        return row;
    }

    public String getName() {
        return name;
    }

    public String getReference() {
        return reference;
    }

    public String getSource_req() {
        return source_req;
    }

    public void setSource_req(String source_req) {
        this.source_req = source_req;
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
        if (cells > 5) new_req = safeLoadString(xrow, 5); // New requirement flag
        if (cells > 6) integration = safeLoadString(xrow, 6); // Integration requirement
        if (cells > 7) service = safeLoadString(xrow, 7); // Integration service requirement
        if (cells > 8) comment = safeLoadString(xrow, 8); // Comment for requirement
        if (cells > 9) linked = safeLoadString(xrow,9); // Linked requirement
        if (cells > 10) curr_status = safeLoadString(xrow,10); // Requirement status
        if (cells > 11) type = safeLoadString(xrow,11); // Requirement type
        if (cells > 12) source = safeLoadString(xrow,12); // Requirement source
        if (cells > 13) foundation = safeLoadString(xrow,13); // Requirement foundation

        if (cells > 14) version = safeLoadString(xrow,14); // Plan to realised in version
        if (cells > 15) release = safeLoadString(xrow, 15); // Plan to realized in release
        if (cells > 16) questions = safeLoadString(xrow,16); // Work questions for requirement
        if (cells > 17) source_req = safeLoadString(xrow,17); // Requirement in source
        if (cells > 18) tt = safeLoadString(xrow,18); // Team track task
        if (cells > 19) trello = safeLoadString(xrow,19); // Trello task
        if (cells > 20) primary = safeLoadString(xrow,20); // Primary responsible
        if (cells > 21) secondary = safeLoadString(xrow,21); // Secondary responsible
        if (cells > 22) risk = safeLoadString(xrow,22); // Risk
        if (cells > 23) risk_desc = safeLoadString(xrow,23); // Risk description

        // Id evaluation - get all parent nodes
        int outlineLevel = xrow.getOutlineLevel();
        if ((outlineLevel + 1)!= level) {
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
     * Fills XLSX row from object
     * @param row - XLSX row
     */
    public void saveToRow(XSSFRow row) {
        XSSFCell cell = row.createCell(0); cell.setCellValue(level); // Requirement level
        cell = row.createCell(1); cell.setCellValue(name); // Requirement
        cell = row.createCell(2); cell.setCellValue(priority); // Requirement priority
        cell = row.createCell(3); cell.setCellValue(done); // Requirement has realised
        cell = row.createCell(4); cell.setCellValue(reference); // Requirement from other source (MarxWeb)
        cell = row.createCell(5); cell.setCellValue(new_req); // New requirement flag
        cell = row.createCell(6); cell.setCellValue(integration); // Integration requirement
        cell = row.createCell(7); cell.setCellValue(service); // Integration service requirement
        cell = row.createCell(8); cell.setCellValue(comment); // Comment for requirement
        cell = row.createCell(9); cell.setCellValue(linked); // Linked requirement
        cell = row.createCell(10); cell.setCellValue(curr_status); // Requirement status
        cell = row.createCell(11); cell.setCellValue(type); // Requirement type
        cell = row.createCell(12); cell.setCellValue(source); // Requirement source
        cell = row.createCell(13); cell.setCellValue(foundation); // Requirement foundation

        cell = row.createCell(14); cell.setCellValue(version); // Plan to realised in version
        cell = row.createCell(15); cell.setCellValue(release); // Plan to realized in release
        cell = row.createCell(16); cell.setCellValue(questions); // Work questions for requirement
        cell = row.createCell(17); cell.setCellValue(source_req); // Source requirement
        cell = row.createCell(18); cell.setCellValue(tt); // Team track task
        cell = row.createCell(19); cell.setCellValue(trello); // Trello task
        cell = row.createCell(20); cell.setCellValue(primary); // Primary responsibly
        cell = row.createCell(21); cell.setCellValue(secondary); // Secondary responsibly
        cell = row.createCell(22); cell.setCellValue(risk); // Risk
        cell = row.createCell(23); cell.setCellValue(risk_desc); // Risk description

    }

    /**
     * Function compares two requirements (this and another), return details: list of indexes for all different columns
     * @param o - requiremet for compare
     * @return null (if something gone wrong) or list of indexes for all different columns (e.g. 0,2,3,5): indexes start from "0"
     */
    public List<Integer> compare(Object o) {

        ArrayList<Integer> changes = new ArrayList<>();

        if (this == o) return null;
        if (o == null || getClass() != o.getClass()) return null;

        Requirement that = (Requirement) o;

        if (!id.equals(that.id)) return null;

        if (level != null ? !level.equals(that.level) : that.level != null) changes.add(0);
        if (name != null ? !name.equals(that.name) : that.name != null) changes.add(1);
        if (priority != null ? !priority.equals(that.priority) : that.priority != null) changes.add(2);
        if (done != null ? !done.equals(that.done) : that.done != null) changes.add(3);
        if (reference != null ? !reference.equals(that.reference) : that.reference != null) changes.add(4);
        if (new_req != null ? !new_req.equals(that.new_req) : that.new_req != null) changes.add(5);
        if (integration != null ? !integration.equals(that.integration) : that.integration != null) changes.add(6);
        if (service != null ? !service.equals(that.service) : that.service != null) changes.add(7);
        if (comment != null ? !comment.equals(that.comment) : that.comment != null) changes.add(8);
        if (linked != null ? !linked.equals(that.linked) : that.linked != null) changes.add(9);
        if (curr_status != null ? !curr_status.equals(that.curr_status) : that.curr_status != null) changes.add(10);
        if (type != null ? !type.equals(that.type) : that.type != null) changes.add(11);
        if (source != null ? !source.equals(that.source) : that.source != null) changes.add(12);
        if (foundation != null ? !foundation.equals(that.foundation) : that.foundation != null) changes.add(13);

        if (version != null ? !version.equals(that.version) : that.version != null) changes.add(14);
        if (release != null ? !release.equals(that.release) : that.release != null) changes.add(15);
        if (questions != null ? !questions.equals(that.questions) : that.questions != null) changes.add(16);
        if (source_req != null ? !source_req.equals(that.source_req) : that.source_req != null) changes.add(17);
        if (tt != null ? !tt.equals(that.tt) : that.tt != null) changes.add(18);
        if (trello != null ? !trello.equals(that.trello) : that.trello != null) changes.add(19);
        if (primary != null ? !primary.equals(that.primary) : that.primary != null) changes.add(20);
        if (secondary != null ? !secondary.equals(that.secondary) : that.secondary != null) changes.add(21);
        if (risk != null ? !risk.equals(that.risk) : that.risk != null) changes.add(22);
        if (risk_desc != null ? !risk_desc.equals(that.risk_desc) : that.risk_desc != null) changes.add(23);

        return changes;
    }

    /**
     * Reads one Excel file (first sheet)
     * @param file - file name
     * @return - array with sheet data
     * @throws IOException - may throws file reading errors
     */
    static LinkedHashMap<String, Requirement> readFromExcel(String file) throws IOException {

        LinkedHashMap<String, Requirement> array = new LinkedHashMap<>();

        XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = book.getSheetAt(0);

        int lastRow = sheet.getLastRowNum();

        for (int rowNum = 0; rowNum <= lastRow; rowNum++) {

            if (rowNum <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) break;

            Requirement req = new Requirement();
            req.loadFromRow(row);
            if (array.containsKey(req.id)) {
                System.out.println("ERROR. Row " + (rowNum + 1) + " contains requirement was already loaded before for row " + (array.get(req.id).getRow() + 1));
            }

            int outlineLevel = sheet.getRow(rowNum).getOutlineLevel();
            int specifiedOutlineLevel = (int) sheet.getRow(rowNum).getCell(0).getNumericCellValue();
            if ((outlineLevel + 1) != specifiedOutlineLevel) {
                System.out.println("ERROR in row " + (rowNum + 1) + ". Real row outline level " + (outlineLevel + 1) + " doesn't suite with level has specified in first column: " + specifiedOutlineLevel);
            }

            array.put(req.id, req);

        }

        book.close();

        return array;
    }

}
