package com.github.houbb.iexcel.sax.handler.impl;

import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.sax.handler.SaxRowHandler;
import com.github.houbb.iexcel.sax.handler.SaxRowHandlerContext;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 默认的行处理工具
 * @author binbin.hou
 * date 2018/11/16 14:09
 */
public class DefaultSaxRowHandler implements SaxRowHandler {

    @Override
    public <T> T handle(SaxRowHandlerContext<T> context) {
        try {
            final Class<T> targetClass = context.getTargetClass();
            final List<Object> cellList = context.getCellList();

            Map<Integer, Field> map = context.getIndexFieldMap();
            T instance = targetClass.newInstance();
            for(Integer index : map.keySet()) {
                Field field = map.get(index);
                field.setAccessible(true);
                final Object cellValue = cellList.get(index);
                field.set(instance, cellValue);
            }
            return instance;
        } catch (InstantiationException | IllegalAccessException e) {
            throw new ExcelRuntimeException(e);
        }
    }

}
