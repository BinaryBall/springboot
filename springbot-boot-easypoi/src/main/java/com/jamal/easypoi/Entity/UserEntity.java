package com.jamal.easypoi.Entity;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.util.Date;

/**
 * 用户实体类
 */
public class UserEntity {

    @Excel(name = "ID")
    private int id;

    @Excel(name = "姓名")
    private String name;

    @Excel(name = "电子邮件",width = 20)
    private String email;

    @Excel(name = "年龄")
    private int age;

    @Excel(name = "性别",replace={"男_1", "女_2"})
    private int sex;

    @Excel(name = "操作时间",format="yyyy-MM-dd HH:mm:ss",width = 20)
    private Date time;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public int getSex() {
        return sex;
    }

    public void setSex(int sex) {
        this.sex = sex;
    }

    public Date getTime() {
        return time;
    }

    public void setTime(Date time) {
        this.time = time;
    }

    @Override
    public String toString() {
        return "UserEntity{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", email='" + email + '\'' +
                ", age=" + age +
                ", sex=" + sex +
                ", time=" + time +
                '}';
    }
}
