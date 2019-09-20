package com.github.houbb.iexcel.util.excel;

import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.core.reader.impl.ExcelReader;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.core.writer.impl.HSSFExcelWriter;
import com.github.houbb.iexcel.core.writer.impl.SXSSFExcelWriter;
import com.github.houbb.iexcel.core.writer.impl.XSSFExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.sax.Sax03ExcelReader;
import com.github.houbb.iexcel.sax.Sax07ExcelReader;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * EXCEL 工具类
 *
 * @author binbin.hou
 * date 2018/11/14 20:06
 */
public final class ExcelUtil {

    /**
     * 构造器私有化
     */
    private ExcelUtil(){}

    //region writer
    /**
     * 获取 2003 版本的 excel
     * 默认为第一个 sheet
     *
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter get03ExcelWriter() {
        return get03ExcelWriter(null);
    }

    /**
     * 获取 2003 版本的 excel
     *
     * @param sheetName sheet 名称
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter get03ExcelWriter(final String sheetName) {
        return new HSSFExcelWriter(sheetName);
    }



    /**
     * 获取 2007 版本的 excel
     * 默认为第一个 sheet
     *
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter get07ExcelWriter() {
        return get07ExcelWriter(null);
    }

    /**
     * 获取 2007 版本的 excel
     *
     *
     * @param sheetName sheet 名称
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter get07ExcelWriter(final String sheetName) {
        return new XSSFExcelWriter(sheetName);
    }


    /**
     * 专门用来写大文件的 excel
     *
     * 默认为第一个 sheet
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter getBigExcelWriter() {
        return getBigExcelWriter(null);
    }

    /**
     * 专门用来写大文件的 excel
     *
     * @param sheetName sheet 名称
     * @return 对应的 excel 生成器
     */
    public static IExcelWriter getBigExcelWriter(final String sheetName) {
        return new SXSSFExcelWriter(sheetName);
    }

    /**
     * 一次写入并且刷新结果
     *
     * @param excelWriter excel 写入
     * @param beanList    bean list
     * @param filePath    文件路径
     */
    public static void onceWriteAndFlush(IExcelWriter excelWriter,
                                         final List<?> beanList,
                                         final String filePath) {
        try (OutputStream outputStream = new FileOutputStream(filePath)) {
            excelWriter.write(beanList);
            excelWriter.flush(outputStream);
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        } finally {
            if (excelWriter != null) {
                try {
                    excelWriter.close();
                } catch (IOException e) {
                    //do nothing...
                }
            }
        }
    }
    //endregion

    //region 读取
    /**
     * 获取 excel 的信息(默认读取第一个 sheet)
     *
     * @param excelFile excel 文件
     * @return excel 普通读取
     */
    public static IExcelReader getExcelReader(final File excelFile) {
        return new ExcelReader(excelFile);
    }

    /**
     * 获取 excel 的信息
     *
     * @param excelFile excel 文件
     * @param sheetIndex sheet 的下标，从0开始
     * @return excel 普通读取
     */
    public static IExcelReader getExcelReader(final File excelFile, final int sheetIndex) {
        return new ExcelReader(excelFile, sheetIndex);
    }

    /**
     * 获取大文件 excel 读取对象(默认读取第一个 sheet)
     *
     * @param excelFile excel 文件信息
     * @return excel 读取类
     */
    public static IExcelReader getBigExcelReader(final File excelFile) {
        final String fileName = excelFile.getName();
        if (fileName.endsWith(ExcelTypeEnum.XLS.getValue())) {
            return new Sax03ExcelReader(excelFile);
        }
        return new Sax07ExcelReader(excelFile);
    }

    /**
     * 获取大文件 excel 读取对象
     *
     * @param excelFile excel 文件信息
     * @param sheetIndex sheet 的下标，从0开始
     * @return excel 读取类
     */
    public static IExcelReader getBigExcelReader(final File excelFile, final int sheetIndex) {
        final String fileName = excelFile.getName();
        if (fileName.endsWith(ExcelTypeEnum.XLS.getValue())) {
            return new Sax03ExcelReader(excelFile, sheetIndex);
        }
        return new Sax07ExcelReader(excelFile, sheetIndex);
    }
    //endregion

}
