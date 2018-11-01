package com.github.houbb.iexcel.constant.enums;

import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.IOException;
import java.io.InputStream;

/**
 *
 * @author jipengfei
 */
public enum ExcelTypeEnum {
    /**
     * 2003 excel
     */
    XLS(".xls"),

    /**
     * 2007 excel
     */
    XLSX(".xlsx"),
    ;

    /**
     * 值
     */
    private final String value;

    private ExcelTypeEnum(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }

    public static ExcelTypeEnum valueOf(InputStream inputStream){
        try {
            FileMagic fileMagic =  FileMagic.valueOf(inputStream);
            if(FileMagic.OLE2.equals(fileMagic)){
                return XLS;
            }
            if(FileMagic.OOXML.equals(fileMagic)){
                return XLSX;
            }
            // i18n
            throw new IllegalArgumentException("excelTypeEnum can not null");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
