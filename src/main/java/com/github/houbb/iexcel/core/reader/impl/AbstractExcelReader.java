package com.github.houbb.iexcel.core.reader.impl;

import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.sax.handler.SaxRowHandler;
import com.github.houbb.iexcel.sax.handler.impl.DefaultSaxRowHandler;

import java.io.File;

/**
 * @author binbin.hou
 * @date 2018/11/16 14:01
 */
public abstract class AbstractExcelReader<T> implements IExcelReader<T> {

    //region 私有变量
    /**
     * 当前 sheet 的下标志
     */
    protected int sheetIndex;

    /**
     * excel 文件信息
     */
    protected File excelFile;

    /**
     * 包含表头
     * 默认第一行是表头(暂时不考虑可以指定哪一行是表头的情况，没有太大实用价值)
     * 如果没有表头可以设置为 false 说明
     */
    protected boolean containsHead = true;

    /**
     * 每一行的处理器
     */
    protected SaxRowHandler saxRowHandler;
    //endregion

    //region 对象初始化
    public AbstractExcelReader(File excelFile) {
        this(excelFile, ExcelConst.DEFAULT_SHEET_INDEX);
    }

    public AbstractExcelReader(File excelFile, int sheetIndex) {
        this.sheetIndex = sheetIndex;
        this.excelFile = excelFile;
        saxRowHandler = new DefaultSaxRowHandler();
    }

    /**
     * 设置是否包含表头
     * @param containsHead  是否包含表头
     * @return this
     */
    public AbstractExcelReader containsHead(final boolean containsHead) {
        this.containsHead = containsHead;
        return this;
    }

    /**
     * 设置是否包含表头
     * @param saxRowHandler 每一行的处理
     * @return this
     */
    public AbstractExcelReader saxRowHandler(final SaxRowHandler saxRowHandler) {
        this.saxRowHandler = saxRowHandler;
        return this;
    }
    //endregion 对象初始化

}
