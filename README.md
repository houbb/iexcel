# 项目简介

本项目用于读取和写入 excel。避免 oom，方便操作。

[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.houbb/iexcel/badge.svg)](http://mvnrepository.com/artifact/com.github.houbb/iexcel)
[![Build Status](https://www.travis-ci.org/houbb/iexcel.svg?branch=master)](https://www.travis-ci.org/houbb/iexcel?branch=master)
[![Coverage Status](https://coveralls.io/repos/github/houbb/iexcel/badge.svg?branch=master)](https://coveralls.io/github/houbb/iexcel?branch=master)

# 创作缘由

实际工作和学习中，apache poi 操作 excel 过于复杂。

近期也看了一些其他的工具框架：

- easypoi

- easyexcel

- hutool-poi

都或多或少难以满足自己的实际需要。

于是就自己写了一个操作 excel 导出的工具。

# 快速开始

## 引入 Jar

使用 maven 管理。

```xml
<dependency>
     <groupId>com.github.houbb</groupId>
     <artifactId>iexcel</artifactId>
     <version>0.0.2</version>
</dependency>
```

## 定义对象

你可以直接参考 [ExcelUtilTest.java](src\test\java\com\github\houbb\iexcel\test\util\ExcelUtilTest.java)

定义一个需要写入/读取的 excel 对象。

- ExcelFieldModel.java

只有声明了 `@ExcelField` 的属性才会被处理，使用说明：[`@ExcelField`](#ExcelField-注解说明)

```java
public class ExcelFieldModel {

    @ExcelField
    private String name;

    @ExcelField(headName = "年龄")
    private String age;

    @ExcelField(mapKey = "EMAIL", writeRequire = false, readRequire = false)
    private String email;

    @ExcelField(mapKey = "ADDRESS", headName = "地址", writeRequire = true)
    private String address;
    
    //getter and setter
}
```

## 写入例子

### IExcelWriter 的实现

IExcelWriter 有几个实现类，你可以直接 new 或者借助 `ExcelUtil` 类去创建。

| IExcelWriter 实现类 | ExcelUtil 如何创建 | 说明 | 
|:---|:---|:---|
| HSSFExcelWriter | ExcelUtil.get03ExcelWriter() | 2003 版本的 excel |
| XSSFExcelWriter | ExcelUtil.get07ExcelWriter() | 2007 版本的 excel |
| SXSSFExcelWriter | ExcelUtil.getBigExcelWriter() | 大文件 excel，避免 OOM |

> [IExcelWriter 接口说明](#IExcelWriter-接口说明)

### 写入到 2003 

- excelWriter03Test()

一个将对象列表写入 2003 excel 文件的例子。

```java
/**
 * 写入到 03 excel 文件
 */
@Test
public void excelWriter03Test() {
    // 待生成的 excel 文件路径
    final String filePath = "excelWriter03.xls";

    // 对象列表
    List<ExcelFieldModel> models = buildModelList();

    try(IExcelWriter excelWriter = ExcelUtil.get03ExcelWriter();
        OutputStream outputStream = new FileOutputStream(filePath)) {
        // 可根据实际需要，多次写入列表
        excelWriter.write(models);

        // 将列表内容真正的输出到 excel 文件
        excelWriter.flush(outputStream);
    } catch (IOException e) {
        throw new ExcelRuntimeException(e);
    }
}
```

- buildModelList()

```java
/**
 * 构建测试的对象列表
 * @return 对象列表
 */
private List<ExcelFieldModel> buildModelList() {
    List<ExcelFieldModel> models = new ArrayList<>();
    ExcelFieldModel model = new ExcelFieldModel();
    model.setName("测试1号");
    model.setAge("25");
    model.setEmail("123@gmail.com");
    model.setAddress("贝克街23号");

    ExcelFieldModel modelTwo = new ExcelFieldModel();
    modelTwo.setName("测试2号");
    modelTwo.setAge("30");
    modelTwo.setEmail("125@gmail.com");
    modelTwo.setAddress("贝克街26号");

    models.add(model);
    models.add(modelTwo);
    return models;
}
```

### 一次性写入到 2007 excel 

有时候列表只写入一次很常见，所有就简单的封装了下：

```java
/**
 * 只写入一次列表
 * 其实是对原来方法的简单封装
 */
@Test
public void onceWriterAndFlush07Test() {
    // 待生成的 excel 文件路径
    final String filePath = "onceWriterAndFlush07.xlsx";

    // 对象列表
    List<ExcelFieldModel> models = buildModelList();

    // 对应的 excel 写入对象
    IExcelWriter excelWriter = ExcelUtil.get07ExcelWriter();

    // 只写入一次列表
    ExcelUtil.onceWriteAndFlush(excelWriter, models, filePath);
}
```

## 读取例子

excel 读取时会根据文件名称判断是哪个版本的 excel。 

### IExcelReader 的实现

IExcelReader 有几个实现类，你可以直接 new 或者借助 `ExcelUtil` 类去创建。

| IExcelReader 实现类 | ExcelUtil 如何创建 | 说明 | 
|:---|:---|:---|
| ExcelReader | ExcelUtil.getExcelReader() | 小文件的 excel 读取实现 |
| Sax03ExcelReader | ExcelUtil.getBigExcelReader() | 大文件的 2003 excel 读取实现 |
| Sax07ExcelReader | ExcelUtil.getBigExcelReader() | 大文件的 2007 excel 读取实现 |

> [IExcelReader 接口说明](#IExcelReader-接口说明)

### excel 读取的例子

```java
/**
 * 读取测试
 */
@Test
public void readWriterTest() {
    File file = new File("excelWriter03.xls");
    IExcelReader<ExcelFieldModel> excelReader = ExcelUtil.getExcelReader(file);
    List<ExcelFieldModel> models = excelReader.readAll(ExcelFieldModel.class);
    System.out.println(models);
}
```

# ExcelField 注解说明

`@ExcelField` 的属性说明如下：

| 属性 | 类型 | 默认值 | 说明 |
|:----|:----|:----|:----|
| mapKey | String | `""` | 仅用于生成的入参为 map 时,会将 map.key 对应的值映射到 bean 上。如果不传：默认使用当前字段名称 | 
| headName | String | `""` | excel 表头字段名称，如果不传：默认使用当前字段名称 | 
| writeRequire | boolean | true | excel 文件是否需要写入此字段 | 
| readRequire | boolean | true | excel 文件是否读取此字段 | 

# IExcelWriter 接口说明

```java
/**
 * 写出数据，本方法只是将数据写入Workbook中的Sheet，并不写出到文件<br>
 * <p>
 * data中元素支持的类型有：
 *  <pre>
 * 1. Bean，既元素为一个Bean，第一个Bean的字段名列表会作为首行，剩下的行为Bean的字段值列表，data表示多行 <br>
 * </pre>
 * @param data 数据
 * @return this
 */
IExcelWriter write(Collection<?> data);

/**
 * 写出数据，本方法只是将数据写入Workbook中的Sheet，并不写出到文件<br>
 *  将 map 按照 targetClass 转换为对象列表
 *  应用场景: 直接 mybatis mapper 查询出的 map 结果，或者其他的构造结果。
 * @param mapList map 集合
 * @param targetClass 目标类型
 * @return this
 */
IExcelWriter write(Collection<Map<String, Object>> mapList, final Class<?> targetClass);

/**
 * 将Excel Workbook刷出到输出流
 *
 * @param outputStream 输出流
 * @return this
 */
IExcelWriter flush(OutputStream outputStream);
```

## 指定 sheet

创建 IExcelWriter 的时候，可以指定 sheet 的下标或者名称。来指定写入的 sheet。

## 是否包含表头

创建 IExcelWriter 的后，可以调用 `excelWriter.containsHead(bool)` 指定是否生成 excel 表头。

# IExcelReader 接口说明

```java
/**
 * 读取当前 sheet 的所有信息
 * @param tClass 对应的 javabean 类型
 * @return 对象列表
 */
List<T> readAll(Class<T> tClass);

/**
 * 读取指定范围内的
 * @param tClass 泛型
 * @param startIndex 开始的行信息(从0开始)
 * @param endIndex 结束的行信息
 * @return 读取的对象列表
 */
List<T> read(Class<T> tClass, final int startIndex, final int endIndex);
```

## 指定 sheet

创建 IExcelReader 的时候，可以指定 sheet 的下标或者名称。来指定读取的 sheet。

注意：大文件 sax 读取模式，只支持指定 sheet 的下标。

## 是否包含表头

创建 IExcelReader 的后，可以调用 `excelReader.containsHead(bool)` 指定是否读取 excel 表头。

# 变更日志

> [变更日志](doc/CHANGELOG.md)

# Bug & Issues

欢迎提出宝贵意见：[Bug & Issues](https://github.com/houbb/iexcel/issues)

# 参与开发

如果你想参与到本项目中(编程，文档，测试，推广 etc)，可以发邮件到 `houbinbin.echo@gmail.com` 参与项目开发。

或者直接提交 [PR](https://github.com/houbb/iexcel/pulls)

