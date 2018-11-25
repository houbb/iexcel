package com.github.houbb.iexcel.test.util;

import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import com.github.houbb.iexcel.util.excel.ExcelUtil;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * excel 工具类测试
 * @author binbin.hou
 * @date 2018/11/25
 */
public class ExcelUtilTest {

    /**
     * 写入到 03 excel 文件
     */
    @Test
    public void excelWriter03Test() {
        // 待生成的 excel 文件路径
        final String filePath = "excelWriter03.xls";

        // 对象列表
        List<ExcelFieldModel> models = buildModelList();

        try(IExcelWriter excelWriter = ExcelUtil.get03ExcelWriter();
            OutputStream outputStream = new FileOutputStream(filePath)) {
            // 可根据实际需要，多次写入列表
            excelWriter.write(models);

            // 将列表内容真正的输出到 excel 文件
            excelWriter.flush(outputStream);
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    /**
     * 只写入一次列表
     * 其实是对原来方法的简单封装
     */
    @Test
    public void onceWriterAndFlush07Test() {
        // 待生成的 excel 文件路径
        final String filePath = "onceWriterAndFlush07.xlsx";

        // 对象列表
        List<ExcelFieldModel> models = buildModelList();

        // 对应的 excel 写入对象
        IExcelWriter excelWriter = ExcelUtil.get07ExcelWriter();

        // 只写入一次列表
        ExcelUtil.onceWriteAndFlush(excelWriter, models, filePath);
    }

    /**
     * 读取测试
     */
    @Test
    public void readWriterTest() {
        File file = new File("excelWriter03.xls");
        IExcelReader<ExcelFieldModel> excelReader = ExcelUtil.getExcelReader(file);
        List<ExcelFieldModel> models = excelReader.readAll(ExcelFieldModel.class);
        System.out.println(models);
    }

    /**
     * 构建测试的对象列表
     * @return 对象列表
     */
    private List<ExcelFieldModel> buildModelList() {
        List<ExcelFieldModel> models = new ArrayList<>();
        ExcelFieldModel model = new ExcelFieldModel();
        model.setName("测试1号");
        model.setAge("25");
        model.setEmail("123@gmail.com");
        model.setAddress("贝克街23号");

        ExcelFieldModel modelTwo = new ExcelFieldModel();
        modelTwo.setName("测试2号");
        modelTwo.setAge("30");
        modelTwo.setEmail("125@gmail.com");
        modelTwo.setAddress("贝克街26号");

        models.add(model);
        models.add(modelTwo);
        return models;
    }

}
