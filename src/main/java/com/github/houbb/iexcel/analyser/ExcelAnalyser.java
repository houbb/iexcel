package com.github.houbb.iexcel.analyser;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.metadata.Sheet;

import java.io.InputStream;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public interface ExcelAnalyser {

    void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
              boolean trim);

    void analysis(Sheet sheetParam);

    void analysis();

    List<Sheet> getSheets();

}
