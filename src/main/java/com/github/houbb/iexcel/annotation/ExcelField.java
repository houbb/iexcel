package com.github.houbb.iexcel.annotation;

import java.lang.annotation.*;

/**
 * Excel 字段注解
 *
 * @author binbin.hou
 * @since 0.0.1
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
     * @deprecated 0.0.4 之后废弃，excel 本身应该专注。对象转换可以交给其他工具。
     */
    @Deprecated
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

    /**
     * 指定顺序编号
     * 解释：默认是按照 Field 信息直接反射处理的，但是有一种情况反射的字段顺序可能会是错乱的。
     * 如进行内存对齐的时候，这个概率虽然很低，但是可以通过指定 order 属性来处理。
     * （1）order 值越小，生成 excel 列越靠前。
     * （2）无注解的字段默认 order=0
     * @return 顺序编号
     * @since 0.0.5
     */
    int order() default 0;

}
