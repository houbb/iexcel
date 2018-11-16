package com.github.houbb.iexcel.constant;

/**
 * EXCEL 相关常量
 *
 * @author binbin.hou
 * @date 2018/11/1 11:02
 */
public interface ExcelConst {

    /**
     * 默认的 sheet 名称
     */
    String DEFAULT_SHEET_NAME = "sheet1";

    /**
     * 默认的 sheet 索引
     */
    int DEFAULT_SHEET_INDEX = 0;

    /**
     * 最大列数限制
     */
    int MAX_COLUMNS_LIMIT = 256;

    /**
     * 每个单元格可输入32767个字符（在编辑栏可显示全部32767个字符），
     * 但只能显示1024个字符
     */
    int MAX_COLUMN_LENGTH_LIMIT = 32767;

    interface Excel03 {
        /**
         * EXCEL 最大行数限制
         */
        int MAX_ROWS_NUM = (1 << 16)-1;
    }

    interface Excel07 {
        /**
         * EXCEL 最大行数限制
         */
        int MAX_ROWS_NUM = (1 << 20)-1;
    }

}
