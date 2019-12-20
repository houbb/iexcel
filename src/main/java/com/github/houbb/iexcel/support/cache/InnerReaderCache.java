package com.github.houbb.iexcel.support.cache;

import com.github.houbb.heaven.annotation.ThreadSafe;
import com.github.houbb.heaven.support.cache.impl.AbstractCache;
import com.github.houbb.heaven.support.cache.impl.ClassFieldListCache;
import com.github.houbb.heaven.util.guava.Guavas;
import com.github.houbb.iexcel.annotation.ExcelField;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * 内部读取类 cache
 * @author binbin.hou
 * @since 0.0.7
 */
@ThreadSafe
public class InnerReaderCache extends AbstractCache<Class, Map<String, Field>> {

    @Override
    protected Map<String, Field> buildValue(Class key) {
        List<Field> fieldList = ClassFieldListCache.getInstance().get(key);
        Map<String, Field> map = Guavas.newHashMap(fieldList.size());

        for (Field field : fieldList) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                boolean readRequire = excelField.readRequire();
                if (readRequire) {
                    final String headName = InnerExcelUtil.getFieldHeadName(excelField, field);
                    map.put(headName, field);
                }
            } else {
                //@since0.0.4 默认使用 fieldName
                final String headName = field.getName();
                map.put(headName, field);
            }
        }
        return map;
    }

}
