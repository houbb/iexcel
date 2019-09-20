# `@ExcelField` 简介

有时候我们需要灵活的指定字段属性，比如对应的 excel 表头字段名称。

比如是否要读写这一行内容。

`@ExcelField` 注解就是为此设计。

# 注解说明

```java
public @interface ExcelField {

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

}
```

# 使用例子

```java
public class UserField {

    @ExcelField(headName = "姓名")
    private String name;

    @ExcelField(headName = "年龄")
    private int age;

}
```

这样生成的 excel 表头就是我们指定的中文。

