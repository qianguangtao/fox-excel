package com.mamba.excel.config;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.StrUtil;
import com.mamba.excel.annotation.ExcelColumn;
import com.mamba.excel.annotation.ExcelSheet;
import lombok.Builder;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:22
 * @description: excel配置信息
 */
@Builder
@Data
public class ExcelConfig {

    /**
     * 获取指定类的Excel列配置列表
     *
     * @param clazz 要获取列配置的类
     * @return 返回指定类的Excel列配置列表
     * @throws IllegalArgumentException 如果ExcelColumn注解的value或index为空，则抛出此异常
     */
    public static List<ColumnConfig> getColumnConfig(Class clazz) {
        List<ColumnConfig> excelColumnList = new ArrayList<>();
        if (clazz == null) {
            return excelColumnList;
        }
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn c = field.getDeclaredAnnotation(ExcelColumn.class);
            if (c != null) {
                if (StrUtil.isBlank(c.value()) || c.index() < 0) {
                    throw new IllegalArgumentException("ExcelColumn注解的value和index不能为空");
                }
                excelColumnList.add(ColumnConfig.builder().header(c.value()).fieldName(field.getName()).note(c.note())
                    .index(c.index()).build());
            }
        }
        if (CollectionUtil.isNotEmpty(excelColumnList)) {
            return excelColumnList.stream().sorted((o1, o2) -> {
                return o1.getIndex() - o2.getIndex();
            }).collect(Collectors.toList());
        }
        return excelColumnList;
    }

    /**
     * 获取指定类的Excel表格配置
     *
     * @param clazz 要获取配置的类
     * @return 返回指定类的Excel表格配置
     * @throws IllegalArgumentException 如果ExcelSheet注解的value或index为空，则抛出此异常
     */
    public static SheetConfig getSheetConfig(Class clazz) {
        return Optional.ofNullable(clazz.getDeclaredAnnotation(ExcelSheet.class)).map(excelSheet -> {
            ExcelSheet s = (ExcelSheet)excelSheet;
            if (StrUtil.isBlank(s.value()) || s.index() < 0) {
                throw new IllegalArgumentException("ExcelSheet注解的value和index不能为空");
            }
            return SheetConfig.builder().name(s.value()).index(s.index()).build();
        }).orElse(SheetConfig.builder().name("sheet0").index(0).build());
    }

    /**
     * sheet配置信息
     */
    @Builder
    @Data
    public static class SheetConfig {
        /** sheet中文名，导出时候用 */
        private String name;
        /** sheet下标 */
        private Integer index;
    }

    /**
     * excel列配置信息
     */
    @Builder
    @Data
    public static class ColumnConfig {
        /** 表头 */
        private String header;
        /** SheetConfig.clazz对应的属性名 */
        private String fieldName;
        /** 标题备注 */
        private String note;
        /** column下标 */
        private Integer index;
    }
}
