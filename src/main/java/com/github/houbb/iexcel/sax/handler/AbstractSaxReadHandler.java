package com.github.houbb.iexcel.sax.handler;

import com.github.houbb.iexcel.hutool.api.SaxReadHandler;

import java.util.List;

/**
 * 抽象类
 * @param <T> 泛型
 * @since 1.0.0
 */
public abstract class AbstractSaxReadHandler<T> implements SaxReadHandler<T> {

    protected abstract void doHandle(int i, List<Object> list, T t);

    @Override
    public boolean handle(int i, List<Object> list, T t) {
        this.doHandle(i, list, t);

        return false;
    }

}
