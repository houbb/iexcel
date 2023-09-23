package com.github.houbb.iexcel.sax.handler;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * sax 行数据处理的上下文对象
 * @author binbin.hou
 * date 2018/11/16 15:53
 */
public class SaxRowHandlerContext<T> {

    /**
     * 需要转换的 class 类
     */
    private Class<T> targetClass;

    /**
     * 当前行的 index
     */
    private int rowIndex;

    /**
     * cell 的列表信息
     */
    private List<Object> cellList;

    /**
     * cell 的下标和 field 之间的映射关系
     */
    private Map<Integer, Field> indexFieldMap;

    public Class<T> getTargetClass() {
        return targetClass;
    }

    public void setTargetClass(Class<T> targetClass) {
        this.targetClass = targetClass;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public List<Object> getCellList() {
        return cellList;
    }

    public void setCellList(List<Object> cellList) {
        this.cellList = cellList;
    }

    public Map<Integer, Field> getIndexFieldMap() {
        return indexFieldMap;
    }

    public void setIndexFieldMap(Map<Integer, Field> indexFieldMap) {
        this.indexFieldMap = indexFieldMap;
    }

}
