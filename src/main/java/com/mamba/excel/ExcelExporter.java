package com.mamba.excel;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.convert.Convert;
import cn.hutool.core.date.DatePattern;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.mamba.excel.config.ExcelConfig;
import com.mamba.excel.kit.ExcelKit;
import com.mamba.excel.kit.ExcelSheetData;
import com.mamba.serializer.EnumDefinition;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.format.annotation.DateTimeFormat;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Optional;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:52
 * @description: excel 导出封装类
 */
@Slf4j
public class ExcelExporter {

    /** excel header下标 */
    private static final int HEADER_INDEX = 0;
    /** excel中头的行数，默认1行 */
    private static final int HEADER_ROW_NUMBER = 1;
    /** 表头索引 */
    @Setter
    private int headerIndex;
    /** excel中表头的行数 */
    @Setter
    private int headerRowNumber;
    /** 数据导出工具类 */
    @Getter
    private final ExcelWriter writer;

    /**
     * Excel导出器构造函数。
     *
     * @param headerIndex 表头所在的行索引，从0开始计数。
     * @param headerRowNumber 表头所在的行数，从0开始计数。
     */
    public ExcelExporter(int headerIndex, int headerRowNumber) {
        this.headerIndex = headerIndex;
        this.headerRowNumber = headerRowNumber;
        this.writer = ExcelUtil.getWriter(true);
    }

    /**
     * Excel导出器的默认构造函数。 使用默认的头索引和头行数创建ExcelWriter对象。
     */
    public ExcelExporter() {
        this.headerIndex = HEADER_INDEX;
        this.headerRowNumber = HEADER_ROW_NUMBER;
        this.writer = ExcelUtil.getWriter(true);
    }

    /**
     * 导出数据到Excel文件并发送给客户端
     *
     * @param excelSheetDataList 需要导出的ExcelSheetData列表
     * @param response HttpServletResponse对象，用于将生成的Excel文件发送给客户端
     * @param fileName 导出的Excel文件的名称
     */
    public void exportData(List<ExcelSheetData> excelSheetDataList, HttpServletResponse response, String fileName) {
        fillData(excelSheetDataList);
        doExport(response, fileName);
    }

    /**
     * 导出数据到指定路径的Excel文件中
     *
     * @param excelSheetDataList 需要导出的ExcelSheetData列表
     * @param filePath 导出的Excel文件的路径
     */
    public void exportData(List<ExcelSheetData> excelSheetDataList, String filePath) {
        fillData(excelSheetDataList);
        doExport(filePath);
    }

    /**
     * 填充Excel数据
     *
     * @param excelSheetDataList ExcelSheetData列表
     */
    private void fillData(List<ExcelSheetData> excelSheetDataList) {
        for (ExcelSheetData excelSheetData : excelSheetDataList) {
            ExcelConfig.SheetConfig sheetConfig = ExcelConfig.getSheetConfig(excelSheetData.getSheetDefinition());
            List<ExcelConfig.ColumnConfig> columnConfigList =
                ExcelConfig.getColumnConfig(excelSheetData.getSheetDefinition());
            writer.setSheet(sheetConfig.getIndex());
            writer.renameSheet(sheetConfig.getIndex(), sheetConfig.getName());
            // 设置标题
            fillHeader(columnConfigList);
            if (CollectionUtil.isNotEmpty(excelSheetData.getData())) {
                fillDropdown(columnConfigList, excelSheetData.getData().size());
                fillContent(columnConfigList, excelSheetData.getData());
                ExcelKit.setAutoSizeColumn(writer.getSheet());
            }
        }
    }

    /**
     * 填充下拉框
     *
     * @param columnConfigList 列配置列表
     * @param dateSize 数据大小
     */
    public void fillDropdown(List<ExcelConfig.ColumnConfig> columnConfigList, int dateSize) {
        for (int i = 0; i < columnConfigList.size(); i++) {
            ExcelConfig.ColumnConfig columnConfig = columnConfigList.get(i);
            if (ObjectUtil.isNotNull(columnConfig.getEnumDefinition().getEnumConstants())) {
                EnumDefinition[] enumConstants = columnConfig.getEnumDefinition().getEnumConstants();
                String[] enumValues = new String[enumConstants.length];
                for (int j = 0; j < enumConstants.length; j++) {
                    String str = Convert.toStr(enumConstants[j].getComment());
                    enumValues[j] = str;
                }
                ExcelKit.setDropdownList(writer.getSheet(), HEADER_ROW_NUMBER, columnConfig.getIndex(),
                        HEADER_ROW_NUMBER + dateSize - 1, columnConfig.getIndex(), enumValues);
            }
        }
    }

    /**
     * 导出异常信息，填充下拉框
     *
     * @param columnConfigList 列配置列表
     * @param row 行索引，从0开始计数
     */
    public void fillDropdownRow(List<ExcelConfig.ColumnConfig> columnConfigList, int row) {
        for (int i = 0; i < columnConfigList.size(); i++) {
            ExcelConfig.ColumnConfig columnConfig = columnConfigList.get(i);
            if (ObjectUtil.isNotNull(columnConfig.getEnumDefinition().getEnumConstants())) {
                EnumDefinition[] enumConstants = columnConfig.getEnumDefinition().getEnumConstants();
                String[] enumValues = new String[enumConstants.length];
                for (int j = 0; j < enumConstants.length; j++) {
                    String str = Convert.toStr(enumConstants[j].getComment());
                    enumValues[j] = str;
                }
                ExcelKit.setDropdownList(writer.getSheet(), row + HEADER_ROW_NUMBER, columnConfig.getIndex(),
                        row + HEADER_ROW_NUMBER, columnConfig.getIndex(), enumValues);
            }
        }
    }

    /**
     * 根据对象字段名获取单元格值
     *
     * @param object 要获取值的对象
     * @param field 字段名
     * @return 返回单元格值，如果字段值为空，则返回null
     */
    private Object getCellValue(Object object, String field) {
        Object value = ReflectUtil.getFieldValue(object, field);
        if (ObjectUtil.isEmpty(value)) {
            return null;
        }
        // 根据类型取出对应类型的值
        Object cellValue = value;
        try {
            Field declaredField = object.getClass().getDeclaredField(field);
            Class<?> fieldType = declaredField.getType();
            if (fieldType == LocalDate.class || fieldType == Date.class || fieldType == LocalDateTime.class) {
                DateTimeFormat format = declaredField.getDeclaredAnnotation(DateTimeFormat.class);
                String pattern =
                    fieldType == LocalDate.class ? DatePattern.NORM_DATE_PATTERN : DatePattern.NORM_DATETIME_PATTERN;
                if (ObjectUtil.isNotNull(format) && StrUtil.isNotBlank(format.pattern())) {
                    pattern = format.pattern();
                }
                cellValue = DateUtil.format(Convert.toDate(value), pattern);
            } else if (fieldType.isEnum()) {
                boolean isEnumDefinition =
                    Arrays.stream(Optional.ofNullable(fieldType.getInterfaces()).orElse(new Class[0]))
                        .anyMatch(cls -> cls == EnumDefinition.class);
                if (isEnumDefinition) {
                    EnumDefinition d = (EnumDefinition)value;
                    cellValue = d.getComment();
                } else {
                    Enum e = (Enum)value;
                    cellValue = e.name();
                }
            }
        } catch (Exception e) {
            log.error("反射处理导出单元格值出错", e);
        }
        return cellValue;
    }

    /**
     * 填充Excel表格的内容
     *
     * @param columnConfigList 列配置列表，包含每个字段的配置信息
     * @param contentList 需要填充的数据列表
     */
    public void fillContent(List<ExcelConfig.ColumnConfig> columnConfigList, List contentList) {
        for (int j = 0; j < contentList.size(); j++) {
            for (int k = 0; k < columnConfigList.size(); k++) {
                writer.writeCellValue(k, j + HEADER_ROW_NUMBER,
                    getCellValue(contentList.get(j), columnConfigList.get(k).getFieldName()));
            }
        }
    }

    /**
     * 填充Excel表格行的内容
     *
     * @param columnConfigList 列配置列表，包含每个字段的配置信息
     * @param content 需要填充的数据对象
     * @param row 当前要填充的行号
     */
    public void fillContent(List<ExcelConfig.ColumnConfig> columnConfigList, Object content, int row) {
        for (int k = 0; k < columnConfigList.size(); k++) {
            writer.writeCellValue(k, row + HEADER_ROW_NUMBER,
                getCellValue(content, columnConfigList.get(k).getFieldName()));
        }
    }

    /**
     * 填充表头
     *
     * @param columnConfigList
     */
    public void fillHeader(List<ExcelConfig.ColumnConfig> columnConfigList) {
        for (int i = 0; i < columnConfigList.size(); i++) {
            writer.writeCellValue(i, HEADER_INDEX, columnConfigList.get(i).getHeader());
            ExcelKit.setCellComment(writer, HEADER_INDEX, i, columnConfigList.get(i).getNote());
        }
    }

    /**
     * 执行Excel导出操作
     *
     * @param response HttpServletResponse对象，用于将生成的Excel文件发送给客户端
     * @param fileName 导出的Excel文件的名称
     * @throws IOException 如果在导出过程中出现输入输出异常
     */
    public void doExport(HttpServletResponse response, String fileName) {
        OutputStream out = null;
        try {
            if (StrUtil.isBlank(fileName)) {
                fileName = IdUtil.fastSimpleUUID() + ".xlsx";
            }
            if (fileName.indexOf(".xlsx") == -1) {
                fileName += ".xlsx";
            }
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            out = response.getOutputStream();
            writer.flush(out, true);
        } catch (IOException e) {
            log.error("导出excel异常", e);
        }
        IoUtil.close(out);
    }

    /**
     * 执行Excel导出操作到指定文件路径
     *
     * @param filePath 要导出Excel文件的路径
     * @throws IOException 如果在导出过程中出现输入输出异常
     */
    public void doExport(String filePath) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            writer.flush(out, true);
        } catch (IOException e) {
            log.error("导出excel异常", e);
        }
        IoUtil.close(out);
    }

}
