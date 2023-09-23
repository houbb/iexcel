package com.github.houbb.iexcel.test.model;


import com.github.houbb.iexcel.hutool.annotation.ExcelField;

import java.util.ArrayList;
import java.util.List;

/**
 * 用户对象-指定顺序
 * @author binbin.hou
 * @since 0.0.5
 */
public class UserFieldOrdered {

    @ExcelField(headerName =  "姓名", order = 1)
    private String name;

    @ExcelField(headerName =  "年龄", order = 2)
    private int age;

    @ExcelField(headerName =  "地址", order = 0)
    private String address;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    @Override
    public String toString() {
        return "UserFieldOrdered{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", address='" + address + '\'' +
                '}';
    }

}
