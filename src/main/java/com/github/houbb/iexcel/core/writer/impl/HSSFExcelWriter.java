package com.github.houbb.iexcel.core.writer.impl;


import com.github.houbb.iexcel.constant.ExcelConst;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 2003 版本写法
 * @author binbin.hou
 * date 2018/11/14 13:56
 */
public class HSSFExcelWriter extends AbstractExcelWriter {

    public HSSFExcelWriter() {
    }

    public HSSFExcelWriter(String sheetName) {
        super(sheetName);
    }

    @Override
    protected Workbook getWorkbook() {
        return new HSSFWorkbook();
    }

    @Override
    protected int getMaxRowNumLimit() {
        return ExcelConst.Excel03.MAX_ROWS_NUM;
    }

}
