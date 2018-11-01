package com.github.houbb.iexcel.core.write;


import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.ExcelWriter;
import com.github.houbb.iexcel.core.inner.ExcelBuilder;
import com.github.houbb.iexcel.core.inner.ExcelBuilderImpl;
import com.github.houbb.iexcel.metadata.Sheet;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 *
 * @author jipengfei
 */
public class ExcelWriterImpl implements ExcelWriter {

    private ExcelBuilder excelBuilder;

    /**
     *
     *
     */
    public ExcelWriterImpl(OutputStream outputStream, ExcelTypeEnum typeEnum) {
        this(outputStream, typeEnum, true);
    }

    /**
     *
     *
     * @param outputStream
     * @param typeEnum
     */
    public ExcelWriterImpl(OutputStream outputStream, ExcelTypeEnum typeEnum, boolean needHead) {
        excelBuilder = new ExcelBuilderImpl();
        excelBuilder.init(null, outputStream, typeEnum, needHead);
    }

    /**
     *
     * @param templateInputStream
     * @param outputStream
     * @param typeEnum
     */
    public ExcelWriterImpl(InputStream templateInputStream, OutputStream outputStream, ExcelTypeEnum typeEnum, boolean needHead) {
        excelBuilder = new ExcelBuilderImpl();
        excelBuilder.init(templateInputStream,outputStream, typeEnum, needHead);
    }

    /**
     *
     * @param data
     * @param sheet
     * @return this
     */
    @Override
    public void write(List<?> data, Sheet sheet) {
        excelBuilder.addContent(data, sheet);
    }

    @Override
    public void finish() {
        excelBuilder.finish();
    }
}
