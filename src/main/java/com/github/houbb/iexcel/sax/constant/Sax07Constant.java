package com.github.houbb.iexcel.sax.constant;

/**
 * sax 07 常量
 * @author binbin.hou
 * date 2018/11/19 18:54
 */
public final class Sax07Constant {

    //列元素
    public static final String C_ELEMENT = "c";
    //列中的v元素
    public static final String V_ELEMENT = "v";
    //列中的t元素
    public static final String T_ELEMENT = "t";
    //行元素
    public static final String ROW_ELEMENT = "row";


    //列中属性r
    public static final String R_ATTR = "r";
    //列中属性值
    public static final String S_ATTR_VALUE = "s";
    //列中属性值
    public static final String T_ATTR_VALUE = "t";

    //sheet r:Id前缀
    public static final String RID_PREFIX = "rId";

    //时间格式化字符串
    public static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    //saxParser
    public static final String CLASS_SAXPARSER = "org.apache.xerces.parsers.SAXParser";

    //填充字符串
    public static final String CELL_FILL_STR = "@";

    //列的最大位数
    public static final int MAX_CELL_BIT = 3;

}
