package com.github.houbb.iexcel.test.bs;

import com.github.houbb.heaven.util.io.FileUtil;
import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.bs.ExcelBs;
import com.github.houbb.iexcel.test.model.User;
import org.junit.Ignore;
import org.junit.Test;

import java.io.InputStream;
import java.util.List;

/**
 * 引导类测试
 * @author binbin.hou
 * @since 0.0.4
 */
@Ignore
public class ExcelBsTest {

    /**
     * 写入到 03 excel 文件
     * 直接将列表内容写入到文件
     */
    @Test
    public void writeTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/user.xls";

        // 对象列表
        List<User> models = User.buildUserList();

        // 直接写入到文件
        ExcelBs.newInstance(filePath).write(models);
    }

    /**
     * 写入到 03 excel 文件
     * （1）append 可以多次将列表写入到 buffer 中，write 才进行实际的写入。
     */
    @Test
    public void appendTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/user.xls";

        // 对象列表
        List<User> models = User.buildUserList();

        // append() 可以多次写入，最后统一写入到文件
        ExcelBs.newInstance(filePath).append(models).write();
    }

    /**
     * 读取 excel 文件中所有信息
     */
    @Test
    public void readTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/user.xls";
        List<User> userList = ExcelBs.newInstance(filePath).read(User.class);
        System.out.println(userList);
    }

    /**
     * 指定开始和结束的下标进行文件的读写。
     */
    @Test
    public void readIndexTest() {
        // 待生成的 excel 文件路径
        final String filePath = PathUtil.getAppTestResourcesPath()+"/user.xls";
        List<User> userList = ExcelBs.newInstance(filePath).read(User.class, 1, 1);
        System.out.println(userList);
    }

    /**
     * 获取对应的输出流
     * 直接将列表内容写入到文件
     * @since 0.0.6
     */
    @Test
    @Ignore
    public void bytesTest() {
        // 获取对应文件流
        byte[] bytes = ExcelBs.newInstance()
                .append(User.buildUserList())
                .bytes();

        // 根据文件内容，自行选择应用场景，到如 web 下载。
        // 此处演示文件创建
        final String filePath = PathUtil.getAppTestResourcesPath()+"/bytes.xls";
        FileUtil.createFile(filePath, bytes);
    }

}
