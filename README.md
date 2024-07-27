# 项目简介

[IExcel](https://github.com/houbb/iexcel) 用于优雅地读取和写入 excel。

避免大 excel 出现 oom，简约而不简单。

[![Build Status](https://travis-ci.com/houbb/iexcel.svg?branch=master)](https://travis-ci.com/houbb/iexcel)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.houbb/iexcel/badge.svg)](http://mvnrepository.com/artifact/com.github.houbb/iexcel)
[![](https://img.shields.io/badge/license-Apache2-FF0080.svg)](https://github.com/houbb/iexcel/blob/master/LICENSE.txt)
[![Open Source Love](https://badges.frapsoft.com/os/v2/open-source.svg?v=103)](https://github.com/houbb/iexcel)

# 特性

- 一行代码搞定一切

- OO 的方式操作 excel，编程更加方便优雅。

- sax 模式读取，SXSS 模式写入。避免 excel 大文件 OOM。

- 基于注解，编程更加灵活。

- 设计简单，注释完整。方便大家学习改造。

- 可根据注解指定表头顺序

- 支持 excel 文件内容 bytes[] 内容获取，便于用户自定义操作。

## 变更日志

> [变更日志](CHANGELOG.md)

## v0.0.9 主要变更

Fixed [@ExcelField注解失效问题](https://github.com/houbb/iexcel/issues/7)

# 创作缘由

实际工作和学习中，apache poi 操作 excel 过于复杂。

近期也看了一些其他的工具框架：

- easypoi

- easyexcel

- hutool-poi

都或多或少难以满足自己的实际需要，于是就自己写了一个操作 excel 导出的工具。

实现：在阿里 [easyexcel](https://github.com/alibaba/easyexcel) 的基础上进行封装，提升使用的简易度。

# 快速开始

## 环境要求

jdk1.8+

maven 3.x

## 引入 jar

使用 maven 管理。

```xml
<dependency>
     <groupId>com.github.houbb</groupId>
     <artifactId>iexcel</artifactId>
     <version>1.0.0</version>
</dependency>
```

## Excel 写入

### 示例

```java
// 基本属性
final String filePath = PathUtil.getAppTestResourcesPath()+"/excelHelper.xls";
List<User> models = User.buildUserList();

// 直接写入到文件
ExcelHelper.write(filePath, models);
```

其中：

- User.java

```java
public class User {

    private String name;

    private int age;

    //fluent getter/setter/toString()
}
```

- buildUserList()

构建对象列表方法如下：

```java
/**
 * 构建用户类表
 * @return 用户列表
 * @since 0.0.4
 */
public static List<User> buildUserList() {
    List<User> users = new ArrayList<>();
    users.add(new User().name("hello").age(20));
    users.add(new User().name("excel").age(19));
    return users;
}
```

### 写入效果

excel 内容生成为：

```
name	age
hello	20
excel	19
```

## Excel 读取

### 示例

```java
final String filePath = PathUtil.getAppTestResourcesPath()+"/excelHelper.xls";
List<User> userList = ExcelHelper.read(filePath, User.class);
```

### 信息

```
[User{name='hello', age=20}, User{name='excel', age=19}]
```

## SAX 读

```java
// 待生成的 excel 文件路径
final String filePath = PathUtil.getAppTestResourcesPath()+"/excelReadBySax.xls";

        AbstractSaxReadHandler<User> saxReadHandler = new AbstractSaxReadHandler<User>() {
            @Override
            protected void doHandle(int i, List<Object> list, User user) {
                System.out.println(user);
            }
        };

ExcelHelper.readBySax(User.class, saxReadHandler, filePath);
```

# 拓展阅读

[Excel Export 踩坑注意点+导出方案设计](https://houbb.github.io/2016/07/19/java-tool-excel-export-design-01-overview)

[基于 hutool 的 EXCEL 优化实现](https://houbb.github.io/2016/07/19/java-tool-excel-hutool-opt-01-intro)

[iexcel-excel 大文件读取和写入，解决 excel OOM 问题-01-入门介绍](https://houbb.github.io/2016/07/19/java-tool-excel-iexcel-01-intro)

[iexcel-excel 大文件读取和写入-02-Excel 引导类简介](https://houbb.github.io/2016/07/19/java-tool-excel-iexcel-02-excelbs)

[iexcel-excel 大文件读取和写入-03-@ExcelField 注解介绍](https://houbb.github.io/2016/07/19/java-tool-excel-iexcel-03-excelField)

[iexcel-excel 大文件读取和写入-04-order 指定列顺序](https://houbb.github.io/2016/07/19/java-tool-excel-iexcel-04-order)

[iexcel-excel 大文件读取和写入-05-file bytes 获取文件字节信息](https://houbb.github.io/2016/07/19/java-tool-excel-iexcel-05-file-bytes)

[Aapche POI java excel 操作工具包入门](https://houbb.github.io/2016/07/19/java-tool-excel-poi-01-intro)

# Bug & Issues

欢迎提出宝贵意见：[Bug & Issues](https://github.com/houbb/iexcel/issues)

# 后期 Road-Map

- [ ] 是否有表头的指定

- [ ] 添加类型转换支持

- [ ] 对于枚举值的注解支持

- [ ] 对于样式的注解支持

- [ ] 多 sheet 支持
