package com.github.houbb.iexcel.test.core;

import com.github.houbb.iexcel.core.reader.impl.ExcelReader;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import com.github.houbb.iexcel.util.excel.ExcelUtil;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @author binbin.hou
 * @date 2018/11/14 21:01
 */
public class ExcelReaderTest {

    @Test
    public void readAllTest() {
        final String path = "3.xlsx";
        List<ExcelFieldModel> modelList = new ExcelReader(new File(path)).readAll(ExcelFieldModel.class);
        System.out.println(modelList);
    }

}
