package com.github.houbb.iexcel.test.model;


import com.github.houbb.iexcel.hutool.annotation.ExcelField;

import java.util.ArrayList;
import java.util.List;

/**
 * excel field model 测试 bean
 * @author houbinbin
 * @since 0.0.1
 */
public class ExcelFieldModel {

    @ExcelField
    private String name;

    @ExcelField(headerName =  "年龄")
    private String age;

    @ExcelField(writeRequire = false, readRequire = false)
    private String email;

    @ExcelField(headerName = "地址", writeRequire = true)
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

    @Override
    public String toString() {
        return "ExcelFieldModel{" +
                "name='" + name + '\'' +
                ", age='" + age + '\'' +
                ", email='" + email + '\'' +
                ", address='" + address + '\'' +
                '}';
    }

    /**
     * 构建测试的对象列表
     * @return 对象列表
     * @since 0.0.4
     */
    public static List<ExcelFieldModel> buildModelList() {
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
}
