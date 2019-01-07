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
 * @since 2018/12/26 16:17
 */
@ExcelProperty()
public class WriteModel2 extends BaseRowModel {

    @ExcelProperty(value = {"account"})
    private String p1;

    private String p2;

    private int p3;

    private long p4;

    private String p5;

    private float p6;

    private BigDecimal p7;

    private Date p8;

    private String p9;

    private double p10;

    public String getP1() {
        return p1;
    }

    public void setP1(String p1) {
        this.p1 = p1;
    }

    public String getP2() {
        return p2;
    }

    public void setP2(String p2) {
        this.p2 = p2;
    }

    public int getP3() {
        return p3;
    }

    public void setP3(int p3) {
        this.p3 = p3;
    }

    public long getP4() {
        return p4;
    }

    public void setP4(long p4) {
        this.p4 = p4;
    }

    public String getP5() {
        return p5;
    }

    public void setP5(String p5) {
        this.p5 = p5;
    }

    public float getP6() {
        return p6;
    }

    public void setP6(float p6) {
        this.p6 = p6;
    }

    public BigDecimal getP7() {
        return p7;
    }

    public void setP7(BigDecimal p7) {
        this.p7 = p7;
    }

    public Date getP8() {
        return p8;
    }

    public void setP8(Date p8) {
        this.p8 = p8;
    }

    public String getP9() {
        return p9;
    }

    public void setP9(String p9) {
        this.p9 = p9;
    }

    public double getP10() {
        return p10;
    }

    public void setP10(double p10) {
        this.p10 = p10;
    }
}
