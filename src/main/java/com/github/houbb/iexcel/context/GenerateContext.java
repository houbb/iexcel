package com.github.houbb.iexcel.context;

import com.github.houbb.iexcel.metadata.ExcelHeadProperty;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;

/**
 * @author jipengfei
 */
public interface GenerateContext {


    /**
     * @return current analysis sheet
     */
    Sheet getCurrentSheet();

    /**
     * @return
     */
    Workbook getWorkbook();

    /**
     * @return
     */
    OutputStream getOutputStream();

    /**
     * @param sheet
     */
    void buildCurrentSheet(com.github.houbb.iexcel.metadata.Sheet sheet);

    /**
     * @param table
     */
    void buildTable(com.github.houbb.iexcel.metadata.Table table);

    /**
     * @return
     */
    ExcelHeadProperty getExcelHeadProperty();

    /**
     *
     * @return
     */
    boolean needHead();

    CellStyle getDefaultCellStyle();

}


