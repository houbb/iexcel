# 项目简介

[IExcel](https://github.com/houbb/iexcel) 用于优雅地读取和写入 excel。

避免大 excel 出现 oom，简约而不简单。。

[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.houbb/iexcel/badge.svg)](http://mvnrepository.com/artifact/com.github.houbb/iexcel)
[![Build Status](https://www.travis-ci.org/houbb/iexcel.svg?branch=master)](https://www.travis-ci.org/houbb/iexcel?branch=master)
[![Coverage Status](https://coveralls.io/repos/github/houbb/iexcel/badge.svg?branch=master)](https://coveralls.io/github/houbb/iexcel?branch=master)

# 特性

- OO 的方式操作 excel，编程更加方便优雅。

- sax 模式读取，SXSS 模式写入。避免 excel 大文件 OOM。

- 基于注解，编程更加灵活。

- 写入可以基于对象列表，也可以基于 Map，实际使用更加方便。

- 设计简单，注释完整。方便大家学习改造。

## 变更日志

> [变更日志](doc/CHANGELOG.md)

## v0.0.4 主要变化

- 引入 ExcelBs 引导类，优化使用体验。

# 创作缘由

实际工作和学习中，apache poi 操作 excel 过于复杂。

近期也看了一些其他的工具框架：

- easypoi

- easyexcel

- hutool-poi

都或多或少难以满足自己的实际需要，于是就自己写了一个操作 excel 导出的工具。

# 快速开始

## 环境要求

jdk1.7+

maven 3.x

## 引入 jar

使用 maven 管理。

```xml
<dependency>
     <groupId>com.github.houbb</groupId>
     <artifactId>iexcel</artifactId>
     <version>0.0.4</version>
</dependency>
```

## Excel 写入

### 示例

```java
/**
 * 写入到 excel 文件
 * 直接将列表内容写入到文件
 */
public void writeTest() {
    // 待生成的 excel 文件路径
    final String filePath = PathUtil.getAppTestResourcesPath()+"/excelWriter03.xls";

    // 对象列表
    List<User> models = User.buildUserList();

    // 直接写入到文件
    ExcelBs.newInstance(filePath).write(models);
}
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
/**
 * 读取 excel 文件中所有信息
 */
public void readTest() {
    // 待生成的 excel 文件路径
    final String filePath = PathUtil.getAppTestResourcesPath()+"/excelWriter03.xls";
    List<User> userList = ExcelBs.newInstance(filePath).read(User.class);
    System.out.println(userList);
}
```

### 信息

```
[User{name='hello', age=20}, User{name='excel', age=19}]
```

# 文档

[01-ExcelBs 引导类使用说明](doc/user/01-ExcelBs引导类讲解.md)

[02-ExcelField 注解指定字段属性]()

# Bug & Issues

欢迎提出宝贵意见：[Bug & Issues](https://github.com/houbb/iexcel/issues)

# 参与开发

如果你想参与到本项目中(编程，文档，测试，推广 etc)，可以发邮件到 `houbinbin.echo@gmail.com` 参与项目开发。

或者直接提交 [PR](https://github.com/houbb/iexcel/pulls)

