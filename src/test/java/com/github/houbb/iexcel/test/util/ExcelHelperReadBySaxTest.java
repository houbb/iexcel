package com.github.houbb.iexcel.test.util;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.sax.handler.AbstractSaxReadHandler;
import com.github.houbb.iexcel.test.model.User;
import com.github.houbb.iexcel.util.ExcelHelper;
import org.apache.poi.ss.formula.functions.T;
import org.junit.Ignore;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * excel 工具类测试
 * @author binbin.hou
 * @since 0.0.8
 */
@Ignore
public class ExcelHelperReadBySaxTest {

    /**
     * 构建用户类表
     * @return 用户列表
     * @since 0.0.4
     */
    public static List<User> buildUserList() {
        List<User> users = new ArrayList<>();

        for(int i = 0; i < 10000; i++) {
            users.add(new User().name("hello").age(20));
        }
        return users;
    }


    /**
     * 写入到 03 excel 文件
     * 直接将列表内容写入到文件
     * @since 0.0.8
     */
    @Test
    public void writeTest() {
        // 基本属性
        final String filePath = PathUtil.getAppTestResourcesPath()+"/excelReadBySax.xls";
        List<User> models = buildUserList();

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
        final String filePath = PathUtil.getAppTestResourcesPath()+"/excelReadBySax.xls";

        AbstractSaxReadHandler<User> saxReadHandler = new AbstractSaxReadHandler<User>() {
            @Override
            protected void doHandle(int i, List<Object> list, User user) {
                System.out.println(user);
            }
        };

        ExcelHelper.readBySax(User.class, saxReadHandler, filePath);
    }

}
