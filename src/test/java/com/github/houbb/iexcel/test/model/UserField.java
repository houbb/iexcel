package com.github.houbb.iexcel.test.model;

import com.github.houbb.iexcel.hutool.annotation.ExcelField;

import java.util.ArrayList;
import java.util.List;

/**
 * 用户对象
 * @author binbin.hou
 * @since 0.0.4
 */
public class UserField {

    @ExcelField(headerName =  "姓名")
    private String name;

    @ExcelField(headerName =  "年龄")
    private int age;

    public String name() {
        return name;
    }

    public UserField name(String name) {
        this.name = name;
        return this;
    }

    public int age() {
        return age;
    }

    public UserField age(int age) {
        this.age = age;
        return this;
    }

    @Override
    public String toString() {
        return "User{" +
                "name='" + name + '\'' +
                ", age=" + age +
                '}';
    }

    /**
     * 构建用户类表
     * @return 用户列表
     * @since 0.0.4
     */
    public static List<UserField> buildUserList() {
        List<UserField> users = new ArrayList<>();
        users.add(new UserField().name("hello").age(20));
        users.add(new UserField().name("excel").age(19));
        return users;
    }

}
