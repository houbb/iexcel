package com.github.houbb.iexcel.util.excel;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.core.writer.impl.HSSFExcelWriter;
import com.github.houbb.iexcel.core.writer.impl.XSSFExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * EXCEL 工具类
 * @author binbin.hou
 * @date 2018/11/14 20:06
 */
public final class ExcelUtil {

    /**
     * 根据 excel 的类型获取对应的 excel 生成器
     * @param excelType excel 类型
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter getExcelWriter(final ExcelTypeEnum excelType) {
        return getExcelWriter(excelType, null);
    }

    /**
     * 根据 excel 的类型获取对应的 excel 生成器
     * @param excelType excel 类型
     * @param sheetName sheet 名称
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter getExcelWriter(final ExcelTypeEnum excelType,
                                              final String sheetName) {
        if(ExcelTypeEnum.XLS.equals(excelType)) {
            return new HSSFExcelWriter(sheetName);
        }
        return new XSSFExcelWriter(sheetName);
    }

    /**
     * 一次写入并且刷新结果
     * @param beanList bean list
     * @param filePath 文件路径
     */
    public static void onceWriteAndFlush(final List<?> beanList,
                                         final String filePath) {
        onceWriteAndFlush(ExcelTypeEnum.XLSX, null, beanList, filePath);
    }

    /**
     * 一次写入并且刷新结果
     * @param excelType excel 文件类型
     * @param sheetName sheet 名称
     * @param beanList bean list
     * @param filePath 文件路径
     */
    public static void onceWriteAndFlush(final ExcelTypeEnum excelType,
                                         final String sheetName,
                                         final List<?> beanList,
                                         final String filePath) {
        try(IExcelWriter excelWriter = getExcelWriter(excelType, sheetName);
            OutputStream outputStream = new FileOutputStream(filePath);) {
            excelWriter.write(beanList);
            excelWriter.flush(outputStream);
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        }
    }

}
