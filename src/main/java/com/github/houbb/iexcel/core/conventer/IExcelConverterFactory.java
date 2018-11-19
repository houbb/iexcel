package com.github.houbb.iexcel.core.conventer;

/**
 * @author binbin.hou
 * date 2018/11/14 21:56
 */
public interface IExcelConverterFactory {

    /**
     * 转换
     * @param converterClass 转换的 class 信息
     * @return 转换后的结果
     */
    IExcelConverter getConverter(final Class<? extends IExcelConverter> converterClass);

}
