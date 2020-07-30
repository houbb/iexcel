package com.github.houbb.iexcel.test.util;

import com.github.houbb.heaven.util.nio.PathUtil;
import com.github.houbb.iexcel.test.model.UserRequire;
import com.github.houbb.iexcel.util.ExcelHelper;
import org.junit.Assert;
import org.junit.Test;

import java.util.List;

/**
 * @author binbin.hou
 * @since 0.0.9
 */
public class ExcelRequireTest {

    @Test
    public void writeRequireTest() {
        UserRequire require = new UserRequire();
        require.setName("ryo");
        require.setPassword("123456");
        require.setAge(20);

        final String path = PathUtil.getAppTestResourcesPath() + "/writeRequire.xls";
        ExcelHelper.write(path, require);

        List<UserRequire> userRequires = ExcelHelper.read(path, UserRequire.class);
        Assert.assertEquals("[UserRequire{name='ryo', password='null', age=null}]", userRequires.toString());
    }

}
