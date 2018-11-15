package com.github.houbb.iexcel.annotation;

import com.github.houbb.iexcel.core.conventer.IExcelConverter;

import java.lang.annotation.*;

/**
 * Excel 类注解
 *
 * @author binbin.hou
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelConverter {

    /**
     * 转换器的类信息
     * @return 转换器
     */
    Class<? extends IExcelConverter> value();

}
