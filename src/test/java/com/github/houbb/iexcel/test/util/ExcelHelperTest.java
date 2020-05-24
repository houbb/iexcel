package com.github.houbb.iexcel.test.util;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.bs.ExcelBs;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.test.model.ExcelFieldModel;
import com.github.houbb.iexcel.test.model.User;
import com.github.houbb.iexcel.util.ExcelHelper;
import com.github.houbb.iexcel.util.excel.ExcelUtil;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * excel 工具类测试
 * @author binbin.hou
 * @since 0.0.8
 */
@Ignore
public class ExcelHelperTest {

    /**
     * 写入到 03 excel 文件
     * 直接将列表内容写入到文件
     * @since 0.0.8
     */
    @Test
    public void writeTest() {
        // 基本属性
        final String filePath = PathUtil.getAppTestResourcesPath()+"/excelHelper.xls";
        List<User> models = User.buildUserList();

        // 直接写入到文件
        ExcelHelper.write(filePath, models);
    }

    /**
     * 读取 excel 文件中所有信息
     * @since 0.0.8
     */
    @Test
    public void readTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/excelHelper.xls";
        List<User> userList = ExcelHelper.read(filePath, User.class);
        System.out.println(userList);
    }

}
