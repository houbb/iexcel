package com.github.houbb.iexcel.analyser.sax;

import com.github.houbb.iexcel.analyser.ExcelAnalyser;
import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.context.AnalysisContext;
import com.github.houbb.iexcel.metadata.Sheet;
import com.github.houbb.iexcel.util.TypeUtil;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public abstract class BaseSaxAnalyser implements ExcelAnalyser {

    protected AnalysisContext analysisContext;

    /**
     * 执行
     */
    public abstract void execute();

    @Override
    public void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom, boolean trim) {
    }

    @Override
    public void analysis(Sheet sheetParam) {
        execute();
    }

    @Override
    public void analysis() {
        execute();
    }

}
