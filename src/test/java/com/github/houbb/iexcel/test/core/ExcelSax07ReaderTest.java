package com.github.houbb.iexcel.test.core;

import com.github.houbb.iexcel.sax.Sax07ExcelReader;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.List;

/**
 * @author binbin.hou
 * date 2018/11/14 21:01
 */
public class ExcelSax07ReaderTest {

    @Test
    public void readAllTest() {
        final String path = "1.xlsx";
        List<ExcelFieldModel> modelList = new Sax07ExcelReader<ExcelFieldModel>(new File(path))
                .readAll(ExcelFieldModel.class);
        System.out.println(modelList);
    }

    /**
     * 直接 poi 生成的 excel 是无法正常解析的，如果保存一下就可以。
     */
    @Test
    public void readAllTest2() {
        final String path = "4.xlsx";
        List<ExcelFieldModel> modelList = new Sax07ExcelReader<ExcelFieldModel>(new File(path))
                .readAll(ExcelFieldModel.class);
        System.out.println(modelList);
    }


}
