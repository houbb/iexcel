package com.github.houbb.iexcel.test.core;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.sax.Sax03ExcelReader;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.util.List;

/**
 * @author binbin.hou
 * date 2018/11/14 21:01
 */
@Ignore
public class ExcelSax03ReaderTest {

    @Test
    public void readAllTest() {
        final String path = PathUtil.getAppTestResourcesPath()+"/excelWriter03.xls";
        List<ExcelFieldModel> modelList = new Sax03ExcelReader(new File(path))
                .readAll(ExcelFieldModel.class);
        System.out.println(modelList);
    }

}
