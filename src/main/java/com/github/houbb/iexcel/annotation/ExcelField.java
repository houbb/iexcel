package com.github.houbb.iexcel.annotation;

import java.lang.annotation.*;

/**
 * Excel 字段注解
 *
 * @author binbin.hou
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {

    /**
     * 仅用于生成的入参为 map 时
     * 会将 map.key 对应的值映射到 bean 上。
     * 如果不传：默认使用当前字段名称
     * @return map 对应的键
     */
    String mapKey() default "";

    /**
     * excel 表头字段名称
     * 如果不传：默认使用当前字段名称
     * @return 字段名称
     */
    String headName() default "";

    /**
     * excel 文件是否需要写入此字段
     *
     * @return 是否需要写入此字段
     */
    boolean writeRequire() default true;

    /**
     * excel 文件是否读取此字段
     * @return 是否读取此字段
     */
    boolean readRequire() default true;

//    /**
//     * 当前字段处于 excel 的第几列。
//     * 这属于和 headName 不同的形式。一个包含表头，一个不包含表头。
//     * 优先使用 headName 如果没有则使用 index?
//     *
//     * @return index
//     */
//    int index() default 0;

}
