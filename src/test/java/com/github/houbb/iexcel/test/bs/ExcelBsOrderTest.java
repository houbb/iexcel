package com.github.houbb.iexcel.test.bs;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.bs.ExcelBs;
import com.github.houbb.iexcel.test.model.User;
import com.github.houbb.iexcel.test.model.UserFieldOrdered;
import org.junit.Assert;
import org.junit.Ignore;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * 引导类测试
 * @author binbin.hou
 * @since 0.0.5
 */
@Ignore
public class ExcelBsOrderTest {

    /**
     * 写入到 03 excel 文件
     * 直接将列表内容写入到文件-有序写入
     * @since 0.0.5
     */
    @Test
    public void writeTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/userOrdered.xls";

        // 对象列表
        List<UserFieldOrdered> models = buildUserList();

        // 直接写入到文件
        ExcelBs.newInstance(filePath).write(models);
    }

    /**
     * 构建用户类表
     * @return 用户列表
     * @since 0.0.5
     */
    private static List<UserFieldOrdered> buildUserList() {
        List<UserFieldOrdered> users = new ArrayList<>();
        UserFieldOrdered one = new UserFieldOrdered();
        one.setName("one");
        one.setAge(10);
        one.setAddress("china");
        users.add(one);
        return users;
    }

    /**
     * 读取 excel 文件中所有信息
     * @since 0.0.5
     */
    @Test
    public void readTest() {
        final String filePath = PathUtil.getAppTestResourcesPath()+"/userOrdered.xls";
        List<UserFieldOrdered> userList = ExcelBs.newInstance(filePath)
                .read(UserFieldOrdered.class);

        Assert.assertEquals("[UserFieldOrdered{name='one', age=10, address='china'}]",
                userList.toString());
    }

}
