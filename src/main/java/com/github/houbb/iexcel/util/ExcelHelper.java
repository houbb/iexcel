package com.github.houbb.iexcel.util;

import com.github.houbb.iexcel.bs.ExcelBs;

import java.util.Collection;
import java.util.Collections;
import java.util.List;

/**
 * Excel 操作工具类
 * <p> project: iexcel-ExcelHelper </p>
 * <p> create on 2020/5/24 19:33 </p>
 *
 * @author binbin.hou
 * @since 0.0.8
 */
public final class ExcelHelper {

    private ExcelHelper(){}

    /**
     * 写入到指定的 excel 文件
     * @param targetPath 目标路径
     * @param collection 集合信息
     * @since 0.0.8
     */
    public static void write(final String targetPath,
                             final Collection<?> collection) {
        ExcelBs.newInstance(targetPath).write(collection);
    }

    /**
     * 写入到指定的 excel 文件
     * @param targetPath 目标路径
     * @param instance 单个对象
     * @since 0.0.9
     */
    public static void write(final String targetPath,
                             final Object instance) {
        List<?> list = Collections.singletonList(instance);
        write(targetPath, list);
    }

    /**
     * 读取目标文件的内容
     * @param targetPath 目标文件
     * @param tClass 目标类
     * @param <T> 泛型
     * @return 结果列表
     * @since 0.0.8
     */
    public static <T> List<T> read(final String targetPath,
                                   final Class<T> tClass) {
        return ExcelBs.newInstance(targetPath).read(tClass);
    }

}
