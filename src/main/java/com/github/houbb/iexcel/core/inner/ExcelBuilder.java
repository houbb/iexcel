package com.github.houbb.iexcel.core.inner;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.metadata.Sheet;
import com.github.houbb.iexcel.metadata.Table;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public interface ExcelBuilder {


    void init(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType, boolean needHead);


    void addContent(List data, int startRow);


    void addContent(List data, Sheet sheetParam);


    void addContent(List data, Sheet sheetParam, Table table);

    void finish();

}
