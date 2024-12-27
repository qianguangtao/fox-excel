package com.mamba.excel.kit;

import lombok.Data;

import java.util.List;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 13:50
 * @description: excel sheet 表头信息 + 数据的封装类
 */
@Data
public class ExcelSheetData<T> {
    /** excel sheet 数据 */
    private List<T> data;
    /** excel sheet 表头信息 */
    private Class<T> sheetDefinition;
}
