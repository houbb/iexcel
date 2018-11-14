# 项目简介

本项目用于

[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.houbb/iexcel/badge.svg)](http://mvnrepository.com/artifact/com.github.houbb/iexcel)
[![Build Status](https://www.travis-ci.org/houbb/iexcel.svg?branch=master)](https://www.travis-ci.org/houbb/iexcel?branch=master)
[![Coverage Status](https://coveralls.io/repos/github/houbb/iexcel/badge.svg?branch=master)](https://coveralls.io/github/houbb/iexcel?branch=master)

# 创作缘由

实际工作和学习中，apache poi 操作 excel 过于复杂。

近期也看了一些其他的工具框架：

- easypoi

- easyexcel

- hutool-poi

都或多或少难以满足自己的实际需要。于是就自己写了一个操作 excel 导出的工具。

暂时只支持生成 excel，后期陆续添加新的功能。


# 快速开始

## jar 引入

本项目使用 Maven 管理 jar

```xml
<dependency>
     <groupId>com.github.houbb</groupId>
     <artifactId>iexcel</artifactId>
     <version>${iexcel 对应版本}</version>
 </dependency>
```

## 快速体验

直接根据 bean 列表生成对应的 excel 

- ExcelFieldModel.java

```java
public class ExcelFieldModel {

    @ExcelField
    private String name;

    @ExcelField(headName = "年龄")
    private String age;

    @ExcelField(mapKey = "EMAIL", excelRequire = false)
    private String email;

    @ExcelField(mapKey = "ADDRESS", headName = "地址", excelRequire = true)
    private String address;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
```

- 测试代码

1. 创建对象列表

2. 生成对应的 excel 列表

```java
@Test
public void onceWriteAndFlushTest() {
    final String path = "D:\\github\\iexcel\\src\\test\\java\\com\\github\\houbb\\iexcel\\test\\3.xlsx";
    List<ExcelFieldModel> modelList = new ArrayList();
    ExcelFieldModel indexModel = new ExcelFieldModel();
    indexModel.setName("你好");
    indexModel.setAge("10");
    indexModel.setAddress("地址");
    indexModel.setEmail("1@qq.com");
    modelList.add(indexModel);

    ExcelUtil.onceWriteAndFlush(modelList, path);
}
```

## ExcelUtil 的介绍 

## @ExcelField 功能介绍

# 变更日志

> [变更日志](doc/CHANGELOG.md)

