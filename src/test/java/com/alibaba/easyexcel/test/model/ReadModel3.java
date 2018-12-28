package com.alibaba.easyexcel.test.model;

import java.math.BigDecimal;
import java.util.Date;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2018/12/26 21:49
 */
@ExcelProperty
public class ReadModel3 extends BaseRowModel {
    private String str;

    private Float ff;

    private Integer mm;

    private BigDecimal money;

    private Long times;

    private Double activityCode;

    @ExcelProperty(format = "yyyy-MM-dd")
    private Date date;

    private String lx;

    private String name;

    private String kk;

    public String getStr() {
        return str;
    }

    public void setStr(String str) {
        this.str = str;
    }

    public Float getFf() {
        return ff;
    }

    public void setFf(Float ff) {
        this.ff = ff;
    }

    public Integer getMm() {
        return mm;
    }

    public void setMm(Integer mm) {
        this.mm = mm;
    }

    public BigDecimal getMoney() {
        return money;
    }

    public void setMoney(BigDecimal money) {
        this.money = money;
    }

    public Long getTimes() {
        return times;
    }

    public void setTimes(Long times) {
        this.times = times;
    }

    public Double getActivityCode() {
        return activityCode;
    }

    public void setActivityCode(Double activityCode) {
        this.activityCode = activityCode;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getLx() {
        return lx;
    }

    public void setLx(String lx) {
        this.lx = lx;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getKk() {
        return kk;
    }

    public void setKk(String kk) {
        this.kk = kk;
    }

    @Override
    public String toString() {
        return "JavaModel2{" +
                "str='" + str + '\'' +
                ", ff=" + ff +
                ", mm=" + mm +
                ", money=" + money +
                ", times=" + times +
                ", activityCode=" + activityCode +
                ", date=" + date +
                ", lx='" + lx + '\'' +
                ", name='" + name + '\'' +
                ", kk='" + kk + '\'' +
                '}';
    }
}
