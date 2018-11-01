package com.github.houbb.iexcel.annotation;

import java.lang.annotation.*;

/**
 * Excel 字段注解
 * @author binbin.hou
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
@Documented
public @interface Excel {

     /**
      * 表头说明
      * @return 表头说明
      */
     String[] value() default {""};

     /**
      * 下标
      * @return 下标
      */
     int index() default -1;

}
