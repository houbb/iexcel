package com.github.houbb.iexcel.test.converter;

import com.github.houbb.iexcel.core.conventer.IExcelConverter;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;

import java.util.Collection;

/**
 * @author binbin.hou
 * @date 2018/11/15 16:25
 */
public class ExcelFieldConverter implements IExcelConverter<ExcelFieldModel> {

    @Override
    public Collection<ExcelFieldModel> write(Collection<ExcelFieldModel> originCollection) {
        for(ExcelFieldModel excelFieldMode : originCollection) {
            excelFieldMode.setName(excelFieldMode.getName()+"转换写入之后的信息");
        }
        return originCollection;
    }

}
