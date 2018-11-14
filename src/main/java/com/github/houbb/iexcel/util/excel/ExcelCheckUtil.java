package com.github.houbb.iexcel.util.excel;

import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;

/**
 * excel 校验工具类
 * @author binbin.hou
 * @date 2018/11/14 19:47
 */
public class ExcelCheckUtil {

    /**
     * 校验列的数量
     * @param size 数量
     */
    public static void checkColumnNum(final int size) {
        if(size <= 0
        || size > ExcelConst.MAX_COLUMNS_LIMIT) {
            throw new ExcelRuntimeException("excel 列数不在正确范围内");
        }
    }

}
