package com.github.houbb.iexcel.core.conventer.impl;

import com.github.houbb.iexcel.core.conventer.IExcelConverter;
import com.github.houbb.iexcel.core.conventer.IExcelConverterFactory;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;

/**
 * @author binbin.hou
 * @date 2018/11/15 9:14
 */
public class DefaultExcelConverterFactory implements IExcelConverterFactory {

    @Override
    public IExcelConverter convert(Class<? extends IExcelConverter> converterClass) {
        try {
            return converterClass.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new ExcelRuntimeException(e);
        }
    }

}
