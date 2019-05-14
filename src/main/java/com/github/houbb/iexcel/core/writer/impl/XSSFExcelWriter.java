package com.github.houbb.iexcel.core.writer.impl;


import com.github.houbb.iexcel.constant.ExcelConst;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 2007 版本写法
 * @author binbin.hou
 * date 2018/11/14 13:56
 * @since 0.0.1
 */
public class XSSFExcelWriter extends AbstractExcelWriter {

    public XSSFExcelWriter() {
    }

    public XSSFExcelWriter(String sheetName) {
        super(sheetName);
    }

    @Override
    protected Workbook getWorkbook() {
        return new XSSFWorkbook();
    }

    @Override
    protected int getMaxRowNumLimit() {
        return ExcelConst.Excel07.MAX_ROWS_NUM;
    }
}
