package com.github.houbb.iexcel.sax.util;

/**
 * 列的字段类型
 * @author binbin.hou
 * date 2018/11/19 16:14
 */
public enum CellDataType {

    /** Boolean类型 */
    BOOL("b"),
    /** 类型错误 */
    ERROR("e"),
    /** 计算结果类型 */
    FORMULA("str"),
    /** 富文本类型 */
    INLINESTR("inlineStr"),
    /** 字符串类型 */
    SSTINDEX("s"),
    /** 数字类型 */
    NUMBER(""),
    /** 日期类型 */
    DATE("m/d/yy"),
    /** 空类型 */
    NULL("");

    /** 属性值 */
    private String name;

    /**
     * 构造
     *
     * @param name 类型属性值
     */
    CellDataType(String name) {
        this.name = name;
    }

    /**
     * 获取对应类型的属性值
     *
     * @return 属性值
     */
    public String getName() {
        return name;
    }

    /**
     * 类型字符串转为枚举
     * @param name 类型字符串
     * @return 类型枚举
     */
    public static CellDataType of(String name) {
        if(BOOL.name.equals(name)) {
            return BOOL;
        }else if(ERROR.name.equals(name)) {
            return ERROR;
        }else if(INLINESTR.name.equals(name)) {
            return INLINESTR;
        }else if(SSTINDEX.name.equals(name)) {
            return SSTINDEX;
        }else if(FORMULA.name.equals(name)) {
            return FORMULA;
        }else {
            return NULL;
        }
    }

}
