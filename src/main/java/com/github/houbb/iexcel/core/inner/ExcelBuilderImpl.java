package com.github.houbb.iexcel.core.inner;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.context.GenerateContext;
import com.github.houbb.iexcel.context.GenerateContextImpl;
import com.github.houbb.iexcel.metadata.ExcelColumnProperty;
import com.github.houbb.iexcel.metadata.Sheet;
import com.github.houbb.iexcel.metadata.Table;
import com.github.houbb.iexcel.util.PoiFileUtil;
import com.github.houbb.iexcel.util.TypeUtil;
import org.apache.commons.beanutils.BeanUtilsBean;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

/**
 * @author jipengfei
 */
public class ExcelBuilderImpl implements ExcelBuilder {

    private GenerateContext context;

    @Override
    public void init(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType, boolean needHead) {
        try {
            //初始化时候创建临时缓存目录，用于规避POI在并发写bug
            PoiFileUtil.createPOIFilesDirectory();

            context = new GenerateContextImpl(templateInputStream, out, excelType, needHead);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void addContent(List data, int startRow) {
        if (data != null && data.size() > 0) {
            int rowNum = context.getCurrentSheet().getLastRowNum();
            if (rowNum == 0) {
                Row row = context.getCurrentSheet().getRow(0);
                if (row == null) {
                    if (context.getExcelHeadProperty() == null || !context.needHead()) {
                        rowNum = -1;
                    }
                }
            }
            if (rowNum < startRow) {
                rowNum = startRow;
            }
            for (int i = 0; i < data.size(); i++) {
                int n = i + rowNum + 1;

                addOneRowOfDataToExcel(data.get(i), n);
            }
        }
    }

    @Override
    public void addContent(List data, Sheet sheetParam) {
        context.buildCurrentSheet(sheetParam);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void addContent(List data, Sheet sheetParam, Table table) {
        context.buildCurrentSheet(sheetParam);
        context.buildTable(table);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void addOneRowOfDataToExcel(List<Object> oneRowData, Row row) {
        if (oneRowData != null && oneRowData.size() > 0) {
            for (int i = 0; i < oneRowData.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(context.getDefaultCellStyle());
                Object cellValue = oneRowData.get(i);
                if (cellValue != null) {
                    if (cellValue instanceof String) {
                        cell.setCellValue((String)cellValue);
                    } else if (cellValue instanceof Integer) {
                        cell.setCellValue(Double.parseDouble(cellValue.toString()));
                    } else if (cellValue instanceof Double) {
                        cell.setCellValue((Double)cellValue);
                    } else if (cellValue instanceof Short) {
                        cell.setCellValue((Double.parseDouble(cellValue.toString())));
                    }
                } else {
                    cell.setCellValue((String)null);
                }
            }
        }
    }

    private void addOneRowOfDataToExcel(Object oneRowData, Row row) {
        int i = 0;
        for (ExcelColumnProperty excelHeadProperty : context.getExcelHeadProperty().getColumnPropertyList()) {
            Cell cell = row.createCell(i);
            // 样式
            String cellValue = null;
            try {
                Object value = BeanUtilsBean.getInstance().getPropertyUtils().getNestedProperty(oneRowData,
                    excelHeadProperty.getField().getName());

                if (value instanceof Date) {
                    cellValue = TypeUtil.formatDate((Date)value, excelHeadProperty.getFormat());
                } else {
                    cellValue = BeanUtilsBean.getInstance().getConvertUtils().convert(value);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (TypeUtil.isNotEmpty(cellValue)) {
                if (TypeUtil.isNum(excelHeadProperty.getField())) {
                    cell.setCellValue(Double.parseDouble(cellValue));
                } else {
                    cell.setCellValue(cellValue);
                }
            } else {
                cell.setCellValue("");
            }
            i++;
        }

    }

    private void addOneRowOfDataToExcel(Object oneRowData, int n) {
        Row row = context.getCurrentSheet().createRow(n);
        if (oneRowData instanceof List) {
            addOneRowOfDataToExcel((List)oneRowData, row);
        } else {
            addOneRowOfDataToExcel(oneRowData, row);
        }
    }
}
