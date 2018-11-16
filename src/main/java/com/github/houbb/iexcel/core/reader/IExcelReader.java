package com.github.houbb.iexcel.core.reader;

import java.util.List;

/**
 *
 * @author binbin.hou
 * @date 2018/11/15 19:57
 */
public interface IExcelReader<T> {

    /**
     * 读取当前 sheet 的所有信息
     * @param tClass 对应的 javabean 类型
     * @return 对象列表
     */
    List<T> readAll(Class<T> tClass);

    /**
     * 读取指定范围内的
     * @param tClass 泛型
     * @param startIndex 开始的行信息
     * @param endIndex 结束的行信息
     * @param <T> 泛型
     * @return 读取的对象列表
     */
    List<T> read(Class<T> tClass, final int startIndex, final int endIndex);

}
