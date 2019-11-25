package com.github.houbb.iexcel.constant;

/**
 * EXCEL 相关常量
 * @author binbin.hou
 * date 2018/11/1 11:02
 * @since 0.0.1     v0.0.3 使用类代替接口
 */
public final class ExcelConst {

    private ExcelConst(){}

    /**
     * 内部 excel 临时文件前缀
     * @since 0.0.6
     */
    public static final String INNER_EXCEL_TEMP_PREFIX = "$INNER_EXCEL_TEMP_PREFIX$";

    /**
     * 默认的 sheet 名称
     */
    public static final String DEFAULT_SHEET_NAME = "sheet1";

    /**
     * 默认的 sheet 索引
     */
    public static final int DEFAULT_SHEET_INDEX = 0;

    /**
     * 最大列数限制
     */
    public static final int MAX_COLUMNS_LIMIT = 256;

    /**
     * 每个单元格可输入32767个字符（在编辑栏可显示全部32767个字符），
     * 但只能显示1024个字符
     */
    public static final int MAX_COLUMN_LENGTH_LIMIT = 32767;

    /**
     * excel 2003
     */
    public static class Excel03 {

        private Excel03(){}

        /**
         * EXCEL 最大行数限制
         */
        public static final int MAX_ROWS_NUM = (1 << 16)-1;
    }

    /**
     * excel 2007
     */
    public static class Excel07 {

        private Excel07(){}

        /**
         * EXCEL 最大行数限制
         */
        public static final int MAX_ROWS_NUM = (1 << 20)-1;
    }

}
