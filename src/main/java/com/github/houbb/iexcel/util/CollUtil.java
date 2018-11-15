package com.github.houbb.iexcel.util;

import java.util.Collection;

/**
 * @author binbin.hou
 * @date 2018/11/15 16:06
 */
public class CollUtil {

    // ---------------------------------------------------------------------- isEmpty
    /**
     * 集合是否为空
     *
     * @param collection 集合
     * @return 是否为空
     */
    public static boolean isEmpty(Collection<?> collection) {
        return collection == null || collection.isEmpty();
    }

    /**
     * 集合是否不为空
     *
     * @param collection 集合
     * @return 是否为空
     */
    public static boolean isNotEmpty(Collection<?> collection) {
        return !isEmpty(collection);
    }

}
