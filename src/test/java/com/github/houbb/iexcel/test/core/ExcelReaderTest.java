package com.github.houbb.iexcel.test.core;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.core.reader.impl.ExcelReader;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.util.List;

/**
 * @author binbin.hou
 * date 2018/11/14 21:01
 * @since 0.0.1
 */
@Ignore
public class ExcelReaderTest {

    @Test
    public void readAllTest() {
        final String path = PathUtil.getAppTestResourcesPath()+"/excelWriter03.xls";
        List<ExcelFieldModel> modelList = new ExcelReader(new File(path)).readAll(ExcelFieldModel.class);
        System.out.println(modelList);
    }

}
