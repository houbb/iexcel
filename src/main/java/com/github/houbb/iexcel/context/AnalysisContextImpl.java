package com.github.houbb.iexcel.context;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.metadata.ExcelHeadProperty;
import com.github.houbb.iexcel.metadata.Sheet;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public class AnalysisContextImpl implements AnalysisContext {

    private Object custom;

    private Sheet currentSheet;

    private ExcelTypeEnum excelType;

    private InputStream inputStream;

    private Integer currentRowNum;

    private Integer totalCount;

    private ExcelHeadProperty excelHeadProperty;

    private boolean trim;

    /**
     * 1904 日期系统中支持的第一天是 1904 年 1 月 1。
     * 当您输入日期时，会将日期转换为 1904 年 1 月 1 以来已用天数表示一个序列号。
     * 例如如果输入 1998，7 月 5 则 Microsoft Excel 会将日期转换为序列号 34519。
     */
    private boolean use1904WindowDate = false;

    @Override
    public void setUse1904WindowDate(boolean use1904WindowDate) {
        this.use1904WindowDate = use1904WindowDate;
    }

    @Override
    public Object getCurrentRowAnalysisResult() {
        return currentRowAnalysisResult;
    }

    @Override
    public void interrupt() {
        throw new ExcelRuntimeException("interrupt error");
    }

    @Override
    public boolean use1904WindowDate() {
        return use1904WindowDate;
    }

    @Override
    public void setCurrentRowAnalysisResult(Object currentRowAnalysisResult) {
        this.currentRowAnalysisResult = currentRowAnalysisResult;
    }

    private Object currentRowAnalysisResult;

    public AnalysisContextImpl(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom, boolean trim) {
        this.custom = custom;
        this.inputStream = inputStream;
        this.excelType = excelTypeEnum;
        this.trim = trim;
    }

    @Override
    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
        if (currentSheet.getClazz() != null) {
            buildExcelHeadProperty(currentSheet.getClazz(), null);
        }
    }

    @Override
    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    @Override
    public Object getCustom() {
        return custom;
    }

    public void setCustom(Object custom) {
        this.custom = custom;
    }

    @Override
    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    @Override
    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    @Override
    public Integer getCurrentRowNum() {
        return this.currentRowNum;
    }

    @Override
    public void setCurrentRowNum(Integer row) {
        this.currentRowNum = row;
    }

    @Override
    public Integer getTotalCount() {
        return totalCount;
    }

    @Override
    public void setTotalCount(Integer totalCount) {
        this.totalCount = totalCount;
    }

    @Override
    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    @Override
    public void buildExcelHeadProperty(Class<?> clazz, List<String> headOneRow) {
        if (this.excelHeadProperty == null && (clazz != null || headOneRow != null)) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, new ArrayList<List<String>>());
        }
        if (this.excelHeadProperty.getHead() == null && headOneRow != null) {
            this.excelHeadProperty.appendOneRow(headOneRow);
        }
    }

    @Override
    public boolean trim() {
        return this.trim;
    }
}
