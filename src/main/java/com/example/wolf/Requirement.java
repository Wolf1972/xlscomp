package com.example.wolf;

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
    Long version; // Plan to realised in version
    Long release; // Plan to realized in release
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
        try {
            int cells = row.getLastCellNum();

            if (cells > 0) level = Math.round(row.getCell(0).getNumericCellValue()); // Requirement level
            if (cells > 1) name = row.getCell(1).getStringCellValue(); // Requirement
            if (cells > 2) priority = row.getCell(2).getStringCellValue(); // Requirement priority
            if (cells > 3) done = row.getCell(3).getStringCellValue(); // Requirement has realised
            if (cells > 4) reference = row.getCell(4).getStringCellValue(); // Requirement from other source (MarxxWeb)

            if (cells > 6) integration = row.getCell(6).getStringCellValue(); // Integration requirement
            if (cells > 7) comment = row.getCell(7).getStringCellValue(); // Comment for requirement
            if (cells > 8) linked = row.getCell(8).getStringCellValue(); // Linked requirement

            if (cells > 13) version = Math.round(row.getCell(13).getNumericCellValue()); // Plan to realised in version
            if (cells > 14) release = Math.round(row.getCell(14).getNumericCellValue()); // Plan to realized in release
            if (cells > 15) questions = row.getCell(15).getStringCellValue(); // Work questions for requirement
        }
        catch (Exception e) {
            System.out.println("Error while loading row " + row.getRowNum());
        }
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
