package com.github.houbb.iexcel.core.writer.impl;


import com.github.houbb.iexcel.constant.ExcelConst;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 专门为写大型 excel 而生
 * @author binbin.hou
 * @date 2018/11/14 13:56
 */
public class SXSSFExcelWriter extends AbstractExcelWriter {

    public SXSSFExcelWriter() {
    }

    public SXSSFExcelWriter(String sheetName) {
        super(sheetName);
    }

    @Override
    protected Workbook getWorkbook() {
        return new SXSSFWorkbook();
    }

    @Override
    protected int getMaxRowNumLimit() {
        return ExcelConst.Excel07.MAX_ROWS_NUM;
    }
}
