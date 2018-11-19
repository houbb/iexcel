
package com.github.houbb.iexcel.util;

import com.github.houbb.iexcel.exception.ExcelRuntimeException;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 方法直接来自 hutool
 * @author binbin.hou
 * date 2018/11/14 18:14
 */
public class BeanUtil {

    public static Map<String, Object> beanToMap(Object bean) {
        try {
            Map<String, Object> map = new LinkedHashMap<>();
            List<Field> fieldList = ClassUtil.getAllFieldList(bean.getClass());

            for (Field field : fieldList) {
                final String fieldName = field.getName();
                final Object fieldValue = field.get(bean);
                map.put(fieldName, fieldValue);
            }
            return map;
        } catch (IllegalAccessException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    /**
     * 判断是否为Bean对象<br>
     * 判定方法是是否存在只有一个参数的setXXX方法
     *
     * @param clazz 待测试类
     * @return 是否为Bean对象
     */
    public static boolean isBean(Class<?> clazz) {
        if (ClassUtil.isNormalClass(clazz)) {
            final Method[] methods = clazz.getMethods();
            for (Method method : methods) {
                if (method.getParameterTypes().length == 1 && method.getName().startsWith("set")) {
                    // 检测包含标准的setXXX方法即视为标准的JavaBean
                    return true;
                }
            }
        }
        return false;
    }

}
