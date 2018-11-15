package com.github.houbb.iexcel.test.model;


import com.github.houbb.iexcel.annotation.ExcelConverter;
import com.github.houbb.iexcel.annotation.ExcelField;
import com.github.houbb.iexcel.test.converter.ExcelFieldConverter;

/**
 * excel field model 测试 bean
 * @author houbinbin
 */
@ExcelConverter(ExcelFieldConverter.class)
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
