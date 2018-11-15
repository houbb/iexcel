package com.github.houbb.iexcel.cache;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author binbin.hou
 * @date 2018/11/15 17:33
 */
public enum BeanFieldCache {

    /**
     * 实例对象
     * 单例的实现方式。
     */
    INSTANCE;

    /**
     *
     */
    private ICache<Class<?>, List<Field>> fieldCache = new ConcurrentCache<>();

    public List<Field> get(final Class<?> clazz) {
        return fieldCache.get(clazz);
    }

    public void put(final Class<?> clazz, final List<Field> fieldList) {
        fieldCache.put(clazz, fieldList);
    }

}
