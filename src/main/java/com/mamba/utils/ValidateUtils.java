package com.mamba.utils;

import cn.hutool.core.util.ReflectUtil;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import java.lang.reflect.Field;
import java.util.Set;

/**
 * hibernate-validator校验工具类
 * <a href="http://docs.jboss.org/hibernate/validator/5.4/reference/en-US/html_single/">参考文档</a>
 */
public class ValidateUtils {

    private static final Validator validator;

    static {
        validator = Validation.buildDefaultValidatorFactory().getValidator();
    }

    /**
     * 校验对象
     *
     * @param object 待校验对象
     * @param groups 待校验的组
     */
    public static void validateEntity(final Object object, final Class<?>... groups) {
        final Set<ConstraintViolation<Object>> constraintViolations = validator.validate(object, groups);
        if (!constraintViolations.isEmpty()) {
            for (final ConstraintViolation<Object> constraint : constraintViolations) {
                throw new IllegalStateException(constraint.getMessage());
            }
        }
    }

    /**
     * 对给定的对象进行验证，并返回验证结果。
     *
     * @param object 要验证的对象
     * @param groups 验证组，用于指定要应用的约束条件集
     * @return 包含所有违反约束条件的集合
     */
    public static Set<ConstraintViolation<Object>> validate(final Object object, final Class<?>... groups) {
        return validator.validate(object, groups);
    }

    /**
     * 验证对象的指定属性是否符合约束条件
     *
     * @param object 要验证的对象
     * @param property 要验证的属性名称
     * @param groups 要验证的约束组，可以传入多个约束组
     * @return 包含所有验证失败的约束违规信息的集合
     */
    public static Set<ConstraintViolation<Object>> validateProperty(final Object object, final String property,
        Class<?>... groups) {
        return validator.validateProperty(object, property, groups);
    }

}
