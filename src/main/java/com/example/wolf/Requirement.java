package com.example.wolf;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class Requirement extends BaseRequirement {

    static final int HEADER_LAST_ROW = 1; // Header row index (starts from 0)

    static RequirementColumnDescriber describer = new RequirementColumnDescriber(); // Map with column indexes

    String id; // Requirement id
    private Integer row; // Excel sheet row num (when requirement loads from Excel sheet)
    private Integer[] parentRows; // List of all parents Excel rows
    // Public columns
    private Integer level; // A(0): Requirement level
    private String name; // B(1): Requirement
    private String priority; // C(2): Requirement priority
    private String done; // D(3): Requirement has realised
    private String other; // E(4): Requirement from other source (mxWeb)
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
    private String other_rel; // R(17): Requirement in source (mxWeb)
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
        if (other != null ? !other.equals(that.other) : that.other != null) return false;
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
        if (other_rel != null ? !other_rel.equals(that.other_rel) : that.other_rel != null) return false;
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
        result = 31 * result + (other != null ? other.hashCode() : 0);
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
        result = 31 * result + (other_rel != null ? other_rel.hashCode() : 0);
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

    public String getOther() {
        return other;
    }

    public String getOtherRel() {
        return other_rel;
    }

    public void setOtherRel(String other_rel) {
        this.other_rel = other_rel;
    }

    /**
     * Fills object fields from XLSX row
     * @param xrow - Excel XLSX row
     */
    void loadFromRow(XSSFRow xrow) {
        int cells = xrow.getLastCellNum();
        row = xrow.getRowNum();

        for (Map.Entry<RequirementColumnType, Integer> item : describer.map.entrySet()) {
            RequirementColumnType rqType = item.getKey();
            Integer column = item.getValue();
            if (column != null && cells > column) {
                switch (rqType) {
                    case RQ_LEVEL:       { level = safeLoadInteger(xrow, column); break; }      // Requirement level
                    case RQ_NAME:        { name = safeLoadString(xrow, column); break; }        // Requirement name
                    case RQ_PRIORITY:    { priority = safeLoadString(xrow, column); break; }    // Requirement priority
                    case RQ_DONE:        { done = safeLoadString(xrow, column); break; }        // Requirement has done
                    case RQ_OTHER:       { other = safeLoadString(xrow, column); break; }       // Requirement from other source (mxWeb)
                    case RQ_NEW_REQ:     { new_req = safeLoadString(xrow, column); break; }     // New requirement flag
                    case RQ_INTEGRATION: { integration = safeLoadString(xrow, column); break; } // Integration requirement
                    case RQ_SERVICE:     { service = safeLoadString(xrow, column); break; }     // Integration service requirement
                    case RQ_COMMENT:     { comment = safeLoadString(xrow, column); break; }     // Comment for requirement
                    case RQ_LINKED:      { linked = safeLoadString(xrow, column); break; }      // Linked requirement
                    case RQ_CURR_STATUS: { curr_status = safeLoadString(xrow, column); break; } // Requirement current status
                    case RQ_TYPE:        { type = safeLoadString(xrow, column); break; }        // Requirement type
                    case RQ_SOURCE:      { source = safeLoadString(xrow, column); break; }      // Requirement source
                    case RQ_FOUNDATION:  { foundation = safeLoadString(xrow, column); break; }  // Requirement foundation

                    case RQ_VERSION:     { version = safeLoadString(xrow, column); break; }     // Plan to realised in version
                    case RQ_RELEASE:     { release = safeLoadString(xrow, column); break; }     // Plan to realised in release
                    case RQ_QUESTIONS:   { questions = safeLoadString(xrow, column); break; }   // Work questions for requirement
                    case RQ_OTHER_REL:  { other_rel = safeLoadString(xrow, column); break; }   // Release in other source (mxWeb)
                    case RQ_TT:          { tt = safeLoadString(xrow, column); break; }          // TeamTrack task
                    case RQ_TRELLO:      { trello = safeLoadString(xrow, column); break; }      // Trello task
                    case RQ_PRIMARY:     { primary = safeLoadString(xrow, column); break; }     // Primary responsible
                    case RQ_SECONDARY:   { secondary = safeLoadString(xrow, column); break; }   // Secondary responsible
                    case RQ_RISK:        { risk = safeLoadString(xrow, column); break; }        // Risk
                    case RQ_RISK_DESC:   { risk_desc = safeLoadString(xrow, column); break; }   // Risk description
                }
            }
        }
        if (level == null) {
            System.out.println("ERROR. Level for row " + (row + 1) + " is empty");
            level = 0;
        }
        if (name == null) {
            System.out.println("ERROR. Name for row " + (row + 1) + " is empty");
            name = "";
        }
        else if (name.contains("\\")) {
            System.out.println("WARNING. Name for row " + (row + 1) + " contains \\");
            name = name.replace("\\", "");
        }

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
        for (Map.Entry<RequirementColumnType, Integer> item : describer.map.entrySet()) {
            RequirementColumnType rqType = item.getKey();
            Integer column = item.getValue();
            if (column != null) {
                XSSFCell cell = row.createCell(column);
                switch (rqType) {
                    case RQ_LEVEL:       { cell.setCellValue(level);  break; }      // Requirement level
                    case RQ_NAME:        { cell.setCellValue(name); break; }        // Requirement name
                    case RQ_PRIORITY:    { cell.setCellValue(priority); break; }    // Requirement priority
                    case RQ_DONE:        { cell.setCellValue(done); break; }        // Requirement has done
                    case RQ_OTHER:       { cell.setCellValue(other); break; }       // Requirement from other source (mxWeb)
                    case RQ_NEW_REQ:     { cell.setCellValue(new_req); break; }     // New requirement flag
                    case RQ_INTEGRATION: { cell.setCellValue(integration); break; } // Integration requirement
                    case RQ_SERVICE:     { cell.setCellValue(service); break; }     // Integration service requirement
                    case RQ_COMMENT:     { cell.setCellValue(comment); break; }     // Comment for requirement
                    case RQ_LINKED:      { cell.setCellValue(linked); break; }      // Linked requirement
                    case RQ_CURR_STATUS: { cell.setCellValue(curr_status); break; } // Requirement current status
                    case RQ_TYPE:        { cell.setCellValue(type); break; }        // Requirement type
                    case RQ_SOURCE:      { cell.setCellValue(source); break; }      // Requirement source
                    case RQ_FOUNDATION:  { cell.setCellValue(foundation); break; }  // Requirement foundation

                    case RQ_VERSION:     { cell.setCellValue(version); break; }     // Plan to realised in version
                    case RQ_RELEASE:     { cell.setCellValue(release); break; }     // Plan to realised in release
                    case RQ_QUESTIONS:   { cell.setCellValue(questions); break; }   // Work questions for requirement
                    case RQ_OTHER_REL:   { cell.setCellValue(other_rel); break; }   // Release in other source (mxWeb)
                    case RQ_TT:          { cell.setCellValue(tt); break; }          // TeamTrack task
                    case RQ_TRELLO:      { cell.setCellValue(trello); break; }      // Trello task
                    case RQ_PRIMARY:     { cell.setCellValue(primary); break; }     // Primary responsible
                    case RQ_SECONDARY:   { cell.setCellValue(secondary); break; }   // Secondary responsible
                    case RQ_RISK:        { cell.setCellValue(risk); break; }        // Risk
                    case RQ_RISK_DESC:   { cell.setCellValue(risk_desc); break; }   // Risk description
                }
            }
        }
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
        if (other != null ? !other.equals(that.other) : that.other != null) changes.add(4);
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
        if (other_rel != null ? !other_rel.equals(that.other_rel) : that.other_rel != null) changes.add(17);
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
