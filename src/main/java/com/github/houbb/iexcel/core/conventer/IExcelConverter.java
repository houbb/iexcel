package com.github.houbb.iexcel.core.conventer;

import java.util.Collection;

/**
 * @author binbin.hou
 * @date 2018/11/14 21:56
 */
public interface IExcelConverter<T> {

    /**
     * 转换
     * @param originCollection 原始的集合
     * @return 转换后的结果
     */
    Collection<T> convert(final Collection<T> originCollection);

}
