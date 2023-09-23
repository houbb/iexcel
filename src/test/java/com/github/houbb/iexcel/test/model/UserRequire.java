package com.github.houbb.iexcel.test.model;

import com.github.houbb.iexcel.hutool.annotation.ExcelField;

/**
 * @author binbin.hou
 * @since 0.0.9
 */
public class UserRequire {

    private String name;

    @ExcelField(writeRequire = false, readRequire = false)
    private String password;

    @ExcelField(writeRequire = true, readRequire = false)
    private Integer age;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return "UserRequire{" +
                "name='" + name + '\'' +
                ", password='" + password + '\'' +
                ", age=" + age +
                '}';
    }

}
