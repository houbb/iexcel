package com.github.houbb.iexcel.metadata;

import java.util.List;

/**
 * @author jipengfei
 */
public class Table {
    /**
     */
    private Class<?> clazz;

    /**
     */
    private List<List<String>> head;

    /**
     */
    private Integer tableNo;

    public Table(Integer tableNo) {
        this.tableNo = tableNo;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public Integer getTableNo() {
        return tableNo;
    }

    public void setTableNo(Integer tableNo) {
        this.tableNo = tableNo;
    }
}
