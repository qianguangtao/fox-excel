package com.mamba.serializer;

import cn.hutool.core.util.StrUtil;
import com.alibaba.fastjson.parser.DefaultJSONParser;
import com.alibaba.fastjson.parser.deserializer.ObjectDeserializer;
import com.alibaba.fastjson.serializer.JSONSerializer;
import com.alibaba.fastjson.serializer.ObjectSerializer;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.util.Objects;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/24 11:12
 * @description: fastjson枚举转换器
 */
@Slf4j
public class EnumConverter implements ObjectSerializer, ObjectDeserializer {

    @Override
    public <T> T deserialze(DefaultJSONParser parser, Type type, Object fieldName) {
        String value = parser.parseObject(String.class);
        if (StrUtil.isNotEmpty(value)) {
            try {
                Class<?> enumClass = Class.forName(type.getTypeName());
                Field field = enumClass.getDeclaredField(EnumDefinition.DISPLAY_FIELD);
                field.setAccessible(true);
                for (Object enumConstant : enumClass.getEnumConstants()) {
                    Object fieldValue = field.get(enumConstant);
                    if (Objects.equals(fieldValue.toString(), value)) {
                        return (T)enumConstant;
                    }
                }
                throw new IllegalStateException("枚举参数异常");
            } catch (Exception e) {
                log.error("枚举参数异常: {}", e);
                throw new IllegalStateException(e);
            }
        }
        return null;
    }

    @Override
    public void write(JSONSerializer serializer, Object object, Object fieldName, Type fieldType, int features)
        throws IOException {
        if (Objects.isNull(object)) {
            serializer.write(null);
            return;
        }
        try {
            Class<?> clazz = Class.forName(fieldType.getTypeName());
            if (!clazz.isEnum()) {
                log.error("当前序列化对象不是枚举: {}:{}", clazz.getSimpleName(), fieldName);
                serializer.write(null);
                return;
            }
            if (!EnumDefinition.class.isAssignableFrom(clazz)) {
                log.error("当前序列化对象未实现EnumDefinition: {}", clazz.getSimpleName());
                serializer.write(null);
                return;
            }
            final EnumDefinition enumComment = (EnumDefinition)object;
            serializer.write(enumComment.getCode());
        } catch (ClassNotFoundException e) {
            log.error("当前枚举类型不存在: {}", fieldType.getTypeName());
            throw new RuntimeException(e);
        }

    }

    @Override
    public int getFastMatchToken() {
        return 0;
    }

}
