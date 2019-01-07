package com.alibaba.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * ExcelProperty
 * the annotation can use in class or field, when use in class ,
 * the index filed will failure and set a default value what the index of the fields,
 * and if the value is empty, it will use the field name make excel head
 *
 * @author wujiaming
 */
@Target({ElementType.FIELD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

    /**
     * it's use to make excel head;
     */
    String[] value() default {""};

    /**
     * the index of excel
     */
    int index() default 99999;

    /**
     * default @see com.alibaba.excel.util.TypeUtil
     * if default is not  meet you can set format
     */
    String format() default "";

    /**
     * Whether to ignore
     */
    boolean ignore() default false;
}
