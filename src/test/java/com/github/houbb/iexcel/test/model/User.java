package com.github.houbb.iexcel.test.model;

import java.util.*;

/**
 * 用户对象
 * @author binbin.hou
 * @since 0.0.4
 */
public class User {

    private String name;

    private int age;

    public String name() {
        return name;
    }

    public User name(String name) {
        this.name = name;
        return this;
    }

    public int age() {
        return age;
    }

    public User age(int age) {
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
    public static List<User> buildUserList() {
        List<User> users = new ArrayList<>();
        users.add(new User().name("hello").age(20));
        users.add(new User().name("excel").age(19));
        return users;
    }

}
