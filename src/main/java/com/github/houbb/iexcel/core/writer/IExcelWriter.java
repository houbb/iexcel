package com.github.houbb.iexcel.core.writer;

import java.io.Closeable;
import java.io.OutputStream;
import java.util.Collection;
import java.util.Map;

/**
 * EXCEL 写文件接口
 * @author binbin.hou
 * date 2018/11/1 10:57
 */
public interface IExcelWriter extends Closeable {

    /**
     * 写出数据，本方法只是将数据写入Workbook中的Sheet，并不写出到文件<br>
     * <p>
     * data中元素支持的类型有：
     *  <pre>
     * 1. Bean，既元素为一个Bean，第一个Bean的字段名列表会作为首行，剩下的行为Bean的字段值列表，data表示多行 <br>
     * </pre>
     * @param data 数据
     * @return this
     */
    IExcelWriter write(Collection<?> data);

    /**
     * 写出数据，本方法只是将数据写入Workbook中的Sheet，并不写出到文件<br>
     *  将 map 按照 targetClass 转换为对象列表
     *  应用场景: 直接 mybatis mapper 查询出的 map 结果，或者其他的构造结果。
     * @param mapList map 集合
     * @param targetClass 目标类型
     * @return this
     */
    IExcelWriter write(Collection<Map<String, Object>> mapList, final Class<?> targetClass);

    /**
     * 将Excel Workbook刷出到输出流
     *
     * @param outputStream 输出流
     * @return this
     */
    IExcelWriter flush(OutputStream outputStream);

}
