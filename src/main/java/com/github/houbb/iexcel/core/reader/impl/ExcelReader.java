package com.github.houbb.iexcel.core.reader.impl;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.util.ClassUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * @author binbin.hou
 * @date 2018/11/15 20:00
 */
public class ExcelReader implements IExcelReader {

    /**
     * excel 文件信息
     */
    private File excelFile;

    /**
     * 当前 sheet 的信息
     */
    private Sheet sheet;

    public ExcelReader(File excelFile, int sheetIndex) {
        Workbook workbook = initWorkbook(excelFile);
        this.sheet = workbook.getSheetAt(sheetIndex);
    }

    public ExcelReader(File excelFile, String sheetName) {
        Workbook workbook = initWorkbook(excelFile);
        this.sheet = workbook.getSheet(sheetName);
    }

    /**
     * 获取 workbook 信息
     * @return workbook
     */
    private Workbook initWorkbook(final File excelFile) {
        final String fileName = excelFile.getName();
        if(fileName.endsWith(ExcelTypeEnum.XLS.getValue())) {
            return new HSSFWorkbook();
        }
        if(fileName.endsWith(ExcelTypeEnum.XLSX.getValue())) {
            return new XSSFWorkbook();
        }
        throw new ExcelRuntimeException("不支持的 excel 文件类型!");
    }


    /**
     * 读取所有的信息
     * 1. 这里的第一行怎么处理 表头的信息？还是按照和导出保持一致？
     * @param tClass
     * @param <T>
     * @return
     */
    @Override
    public <T> List<T> readAll(Class<T> tClass) {
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        try {
            List<Field> fieldList = ClassUtil.getAllFieldList(tClass);
            // 如果没有表头 则默认和字段含有

            //firstRowNum 表头信息处理
            for(int i = firstRowNum+1; i < lastRowNum; i++) {
                Row row = sheet.getRow(i);
                final int firstCellNum = row.getFirstCellNum();
                final int lastCellNum = row.getLastCellNum();
                T instance = tClass.newInstance();

                for(int cIndex = firstCellNum; cIndex < lastCellNum; cIndex++) {
                    // 根据字段处理字段信息
                }
            }

            return null;
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    @Override
    public List<Map<String, Object>> readAll() {
        return null;
    }


}
