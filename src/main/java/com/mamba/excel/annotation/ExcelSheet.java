package com.mamba.excel.annotation;

import java.lang.annotation.*;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:38
 * @description: excel sheet注解
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelSheet {
    /**
     * sheet名称，必填
     */
    String value() default "";

    /**
     * sheet下标，必填，从0开始
     */
    int index() default -1;
}
