package com.github.houbb.iexcel.util.excel;

import com.github.houbb.iexcel.support.style.StyleSet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * 方法直接来自 hutool
 * @author binbin.hou
 * @date 2018/11/14 14:52
 */
public class RowUtil {

    /**
     * 写一行数据
     *
     * @param row 行
     * @param rowData 一行的数据
     * @param styleSet 单元格样式集，包括日期等样式
     * @param isHeader 是否为标题行
     */
    public static void writeRow(Row row, Iterable<?> rowData, StyleSet styleSet, boolean isHeader) {
        int i = 0;
        Cell cell;
        for (Object value : rowData) {
            cell = row.createCell(i);
            CellUtil.setCellValue(cell, value, styleSet, isHeader);
            i++;
        }
    }

}
