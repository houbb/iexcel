package com.github.houbb.iexcel.test.model;


import com.github.houbb.iexcel.annotation.Excel;

/**
 * @author jipengfei
 */
public class ExcelPropertyIndexModel {

    @Excel(value = "姓名" ,index = 0)
    private String name;

    @Excel(value = "年龄",index = 1)
    private String age;

    @Excel(value = "邮箱",index = 2)
    private String email;

    @Excel(value = "地址",index = 3)
    private String address;

    @Excel(value = "性别",index = 4)
    private String sax;

    @Excel(value = {"高度", "世界"},index = 5)
    private String heigh;

    @Excel(value = "备注",index = 6)
    private String last;

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

    public String getSax() {
        return sax;
    }

    public void setSax(String sax) {
        this.sax = sax;
    }

    public String getHeigh() {
        return heigh;
    }

    public void setHeigh(String heigh) {
        this.heigh = heigh;
    }

    public String getLast() {
        return last;
    }

    public void setLast(String last) {
        this.last = last;
    }
}
