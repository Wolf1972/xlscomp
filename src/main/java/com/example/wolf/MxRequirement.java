package com.example.wolf;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashMap;

public class MxRequirement extends BaseRequirement {

    static final int HEADER_LAST_ROW = 0; // Header row index (starts from 0)
    static final int LAST_COMMON_COLUMN = 12; // Last copying column index (starts from 0) - to prevent copying service columns

    Integer id; // Requirement id
    private String block; // Block name
    private String name; // Requirement
    private String priority; // Priority/quality
    private String mxcomment; // MXWeb comment
    private String bankcomment; // Comment
    private String backlog; // Add to backlog
    private Double weight; // Weight
    private String release; // Release
    private String platform; // Platform name
    private String mxwebid; // MXWeb id

    /**
     * Fills object fields from XLSX row
     * @param xrow - Excel XLSX row
     */
    void loadFromRow(XSSFRow xrow) {
        int cells = xrow.getLastCellNum();

        if (cells > 0) id = safeLoadInteger(xrow, 0); // Requirement id
        if (cells > 1) block = safeLoadString(xrow, 1); // Block name
        if (cells > 2) name = safeLoadString(xrow, 2); // Requirement
        if (cells > 3) priority = safeLoadString(xrow, 3); // Priority/quality
        if (cells > 4) mxcomment = safeLoadString(xrow, 4); // MXWeb comment
        if (cells > 5) bankcomment = safeLoadString(xrow, 5); // Comment
        if (cells > 6) backlog = safeLoadString(xrow, 6); // Add to backlog

        if (cells > 8) weight = safeLoadDouble(xrow, 8); // Weight
        if (cells > 9) release = safeLoadString(xrow, 9); // Release
        if (cells > 10) platform = safeLoadString(xrow, 10); // Platform name
        // if (cells > 11) mxwebid = safeLoadString(xrow, 11); // MXWeb id
        if (cells > 11) mxwebid = "M_" + id; // MXWeb id

    }

    @Override
    public String toString() {
        return "MxRequirement{" +
                "id=" + id +
                ", mxwebid='" + mxwebid + '\'' +
                '}';
    }

    public String getMxwebid() {
        return mxwebid;
    }

    public void setMxwebid(String mxwebid) {
        this.mxwebid = mxwebid;
    }

    /**
     * Reads one Excel file (first sheet)
     * @param file - file name
     * @return - array with sheet data
     * @throws IOException - may throws file reading errors
     */
    static LinkedHashMap<Integer, MxRequirement> readFromExcel(String file) throws IOException {

        LinkedHashMap<Integer, MxRequirement> array = new LinkedHashMap<>();

        XSSFWorkbook book = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = book.getSheetAt(0);

        int lastRow = sheet.getLastRowNum();

        for (int rowNum = 0; rowNum <= lastRow; rowNum++) {

            if (rowNum <= Requirement.HEADER_LAST_ROW) continue; // Skip header

            XSSFRow row = sheet.getRow(rowNum);
            if (row == null) break;

            MxRequirement req = new MxRequirement();
            req.loadFromRow(row);
            array.put(req.id, req);
        }

        book.close();

        return array;
    }
}
