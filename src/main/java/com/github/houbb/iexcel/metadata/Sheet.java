package com.github.houbb.iexcel.metadata;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author jipengfei
 */
public class Sheet {

    /**
     */
    @Deprecated
    private int headLineMun;

    /**
     */
    private int sheetNo;

    /**
     */
    private String sheetName;

    /**
     */
    private Class<?> clazz;

    /**
     */
    private List<List<String>> head;

    /**
     * column with
     */
    private Map<Integer,Integer> columnWidthMap = new HashMap<Integer, Integer>();

    /**
     *
     */
    private Boolean autoWidth = Boolean.FALSE;

    /**
     *
     */
    private int startRow = -1;


    public Sheet(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public Sheet(int sheetNo, int headLineMun) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
    }

    public Sheet(int sheetNo, int headLineMun, Class<?> clazz) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
        this.clazz = clazz;
    }

    public Sheet(int sheetNo, int headLineMun, Class<?> clazz, String sheetName,
                 List<List<String>> head) {
        this.sheetNo = sheetNo;
        this.clazz = clazz;
        this.headLineMun = headLineMun;
        this.sheetName = sheetName;
        this.head = head;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
        if (headLineMun == 0) {
            this.headLineMun = 1;
        }
    }

    public int getHeadLineMun() {
        return headLineMun;
    }

    public void setHeadLineMun(int headLineMun) {
        this.headLineMun = headLineMun;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Map<Integer, Integer> getColumnWidthMap() {
        return columnWidthMap;
    }

    public void setColumnWidthMap(Map<Integer, Integer> columnWidthMap) {
        this.columnWidthMap = columnWidthMap;
    }

    public Boolean getAutoWidth() {
        return autoWidth;
    }

    public void setAutoWidth(Boolean autoWidth) {
        this.autoWidth = autoWidth;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    @Override
    public String toString() {
        return "Sheet{" +
                "headLineMun=" + headLineMun +
                ", sheetNo=" + sheetNo +
                ", sheetName='" + sheetName + '\'' +
                ", clazz=" + clazz +
                ", head=" + head +
                ", columnWidthMap=" + columnWidthMap +
                ", autoWidth=" + autoWidth +
                ", startRow=" + startRow +
                '}';
    }
}
