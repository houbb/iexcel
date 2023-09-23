package com.github.houbb.iexcel.core.reader;

import com.github.houbb.iexcel.hutool.api.SaxReadHandler;

import java.util.List;

/**
 * 抽象类
 * @param <T> 泛型
 */
public abstract class AbstractExcelReader<T> implements IExcelReader<T> {

    @Override
    public List<T> readAll(Class<T> tClass) {
        return read(tClass, 0, Integer.MAX_VALUE);
    }

    @Override
    public List<T> read(Class<T> tClass, int startIndex, int endIndex) {
        throw new UnsupportedOperationException();
    }

}
