package com.mamba.serializer;

import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;

@JsonSerialize(using = EnumSerializer.class)
@JsonDeserialize(using = EnumDeserializer.class)
public interface EnumDefinition<T> {
    /**
     * 常量，数据库字段对应枚举类的属性名
     */
    String DB_FIELD = "code";
    /**
     * 常量，枚举显示属性名，一般用于导入导出
     */
    String DISPLAY_FIELD = "comment";

    /**
     * @return 枚举默认的name
     */
    String name();

    /**
     * @return 枚举在数据库存的值（根据系统设计，有可能是0,1,2的code，有可能直接用枚举name）
     */
    T getCode();

    /**
     * @return 枚举的备注，用户返回前端展示
     */
    String getComment();

}
