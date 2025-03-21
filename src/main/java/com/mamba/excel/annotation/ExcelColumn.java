package com.mamba.excel.annotation;

import com.mamba.serializer.EnumDefinition;

import java.lang.annotation.*;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:38
 * @description: excel列注解
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelColumn {
    /**
     * excel列头，必填
     */
    String value() default "";

    /**
     * 单元格备注，可选
     *
     * @return
     */
    String note() default "";

    /**
     * 列下标，必填，从0开始
     */
    int index() default -1;

    /**
     * 枚举类型，可选，用来生成单元格下拉框
     *
     * @return
     */
    Class<? extends EnumDefinition> enumDefinition() default EnumDefinition.class;
}
