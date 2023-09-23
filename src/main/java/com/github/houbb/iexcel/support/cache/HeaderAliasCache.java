package com.github.houbb.iexcel.support.cache;

import com.github.houbb.heaven.annotation.ThreadSafe;
import com.github.houbb.heaven.support.cache.impl.AbstractCache;
import com.github.houbb.heaven.support.cache.impl.ClassFieldListCache;
import com.github.houbb.heaven.support.tuple.impl.Pair;
import com.github.houbb.heaven.util.guava.Guavas;
import com.github.houbb.heaven.util.lang.ObjectUtil;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.hutool.annotation.ExcelField;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.commons.collections4.MapUtils;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/**
 * 头别称 cache
 * @author binbin.hou
 * @since 0.0.7
 */
@ThreadSafe
public class HeaderAliasCache extends AbstractCache<Class, Map<String, String>> {

    @Override
    protected Map<String, String> buildValue(Class beanClass) {
        //cache 获取字段信息
        List<Field> fieldList = ClassFieldListCache.getInstance().get(beanClass);

        // 指定初始化大小，避免扩容
        Map<String, String> resultMap = Guavas.newLinkedHashMap(fieldList.size());

        // 根据 order 进行排序处理
        Map<Integer, List<Pair<String, String>>> orderedMap = new TreeMap<>();

        for (Field field : fieldList) {
            final String fieldName = field.getName();

            // 默认值
            int order = 0;
            Pair<String, String> pair = Pair.of(fieldName, fieldName);

            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excel = field.getAnnotation(ExcelField.class);
                final String headName = InnerExcelUtil.getFieldHeadName(excel, field);
                if (excel.writeRequire()) {
                    pair = Pair.of(fieldName, headName);
                    order = excel.order();
                } else {
                    // 跳过写入
                    //FIXED: https://github.com/houbb/iexcel/issues/7
                    continue;
                }
            }

            // 设置
            List<Pair<String, String>> pairList = orderedMap.get(order);
            if(ObjectUtil.isNull(pairList)) {
                pairList = Guavas.newArrayList();
            }
            pairList.add(pair);
            orderedMap.put(order, pairList);
        }

        // 按照列表循环处理
        if(MapUtils.isEmpty(orderedMap)) {
            throw new ExcelRuntimeException("excel 表头信息为空");
        }

        // 按照顺序构建 map 信息
        for(Map.Entry<Integer, List<Pair<String, String>>> entry : orderedMap.entrySet()) {
            List<Pair<String, String>> pairList = entry.getValue();
            for(Pair<String, String> pair : pairList) {
                resultMap.put(pair.getValueOne(), pair.getValueTwo());
            }
        }
        return resultMap;
    }

}
