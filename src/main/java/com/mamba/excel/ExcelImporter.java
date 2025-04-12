package com.mamba.excel;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.alibaba.fastjson.JSON;
import com.mamba.excel.annotation.ExcelSheet;
import com.mamba.excel.config.ExcelConfig;
import com.mamba.excel.handler.AbstractExcelDataHandler;
import com.mamba.excel.handler.ExcelDataHandlerFactory;
import com.mamba.excel.kit.ExcelKit;
import com.mamba.excel.kit.ImportResultDTO;
import com.mamba.utils.WebUtil;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Supplier;
import java.util.stream.Collectors;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 9:13
 * @description: excel 导入封装类
 */
@Slf4j
public class ExcelImporter {

    /** excel header下标 */
    private static final int HEADER_INDEX = 0;
    /** excel中表头的行数，默认1行 */
    private static final int HEADER_ROW_NUMBER = 1;
    /** 表头索引 */
    @Setter
    private int headerIndex;
    /** excel中表头的行数 */
    @Setter
    private int headerRowNumber;
    /** 保存导入excel中的所有数据 */
    @Getter
    private Map<String, List> allDataMap = new HashMap<>();
    /** 保存导入excel中成功和错误数据的条数 */
    @Getter
    private final ImportResultDTO importResultDTO = new ImportResultDTO();
    /** 是否有错误数据 */
    @Getter
    private boolean hasErrorData = false;
    /** 错误数据导出工具类 */
    private final ExcelExporter errorExcelExporter;
    /** Excel读取工具类 */
    private final ExcelReader reader;
    /** sheet配置 */
    private ExcelConfig.SheetConfig sheetConfig;
    /** 列配置 */
    private List<ExcelConfig.ColumnConfig> columnConfigList;

    /**
     * 构造方法，使用本地磁盘excel初始化ExcelImporter对象。
     *
     * @param filePath Excel文件的路径。
     * @param headerIndex 表头在文件中的索引位置，从0开始。
     * @param headerRowNumber 表头所在的行号，从0开始。
     */
    public ExcelImporter(String filePath, int headerIndex, int headerRowNumber) {
        this.headerIndex = headerIndex;
        this.headerRowNumber = headerRowNumber;
        this.errorExcelExporter = new ExcelExporter(headerIndex, headerRowNumber);
        this.reader = ExcelUtil.getReader(filePath);
    }

    /**
     * 构造方法，使用本地磁盘excel初始化ExcelImporter对象。
     *
     * @param filePath Excel文件的路径
     */
    public ExcelImporter(String filePath) {
        this.headerIndex = HEADER_INDEX;
        this.headerRowNumber = HEADER_ROW_NUMBER;
        this.errorExcelExporter = new ExcelExporter(headerIndex, headerRowNumber);
        this.reader = ExcelUtil.getReader(filePath);
    }

    /**
     * 构造方法，使用web上传的excel初始化ExcelImporter对象。
     *
     * @param file 上传的MultipartFile文件
     * @throws RuntimeException 如果读取文件输入流时发生IO异常，将抛出运行时异常
     */
    public ExcelImporter(MultipartFile file) {
        this.headerIndex = HEADER_INDEX;
        this.headerRowNumber = HEADER_ROW_NUMBER;
        this.errorExcelExporter = new ExcelExporter(headerIndex, headerRowNumber);
        try {
            this.reader = ExcelUtil.getReader(file.getInputStream());
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 导入数据，并将异常数据excel导出到本地磁盘。
     *
     * @param sheetDefinitionList 表格定义列表，每个元素代表一个表格的类定义
     * @param errorExcelPath 要导出的Excel文件的路径
     * @param function 用于处理导入结果回调
     */
    public void importData(List<Class> sheetDefinitionList, String errorExcelPath, BiFunction<ImportResultDTO, ExcelExporter, Boolean> function) {
        importData(sheetDefinitionList);
        Boolean sheetCheck = function.apply(importResultDTO, errorExcelExporter);
        // 如果有异常数据或者自定义的function返回false，则导出异常数据到客户端
        if (hasErrorData || Boolean.FALSE.equals(sheetCheck)) {
            ExcelKit.setAutoSizeColumn(this.errorExcelExporter.getWriter());
            this.errorExcelExporter.doExport(errorExcelPath);
        }
    }

    /**
     * 导入数据，并将异常数据excel通过HttpServletResponse下载。
     *
     * @param sheetDefinitionList 表格定义列表，用于指定每个表格的类定义
     * @param response 响应对象，用于向客户端发送响应
     * @param errorExcelName 导出的异常Excel文件名
     * @param success 无异常数据时候的application/json响应体
     * @param function 用于处理导入结果回调
     */
    public void importData(List<Class> sheetDefinitionList, HttpServletResponse response, String errorExcelName,
        Supplier success, BiFunction<ImportResultDTO, ExcelExporter, Boolean> function) {
        importData(sheetDefinitionList);
        Boolean sheetCheck = function.apply(importResultDTO, errorExcelExporter);
        // 如果有异常数据或者自定义的function返回false，则导出异常数据到客户端
        if (hasErrorData || Boolean.FALSE.equals(sheetCheck)) {
            ExcelKit.setAutoSizeColumn(this.errorExcelExporter.getWriter());
            this.errorExcelExporter.doExport(response, errorExcelName);
        } else {
            WebUtil.writeJson2Response(response, success.get());
        }
    }

    /**
     * 导入数据并处理。
     *
     * @param sheetDefinitionList 表格定义列表，每个元素代表一个表格的类定义
     */
    public void importData(List<Class> sheetDefinitionList) {
        for (Class sheetDefinition : sheetDefinitionList) {
            ExcelSheet excelSheet = (ExcelSheet) sheetDefinition.getAnnotation(ExcelSheet.class);
            AbstractExcelDataHandler excelDataHandler = ExcelDataHandlerFactory.getExcelDataHandler(sheetDefinition);
            sheetConfig = ExcelConfig.getSheetConfig(sheetDefinition);
            columnConfigList = ExcelConfig.getColumnConfig(sheetDefinition);
            List originExcelDataList = getOriginExcelData(sheetDefinition);
            allDataMap.put(sheetDefinition.getName(),
                    CollectionUtil.defaultIfEmpty(originExcelDataList, Collections.emptyList()));
            generateErrorExcelHeader();
            Map<String, Integer> columnConfigMap = columnConfigList.stream()
                    .collect(Collectors.toMap(ExcelConfig.ColumnConfig::getFieldName, ExcelConfig.ColumnConfig::getIndex));
            if (CollectionUtil.isNotEmpty(originExcelDataList)) {
                // 有效数据
                List validDataList = new ArrayList();
                // 无效数据
                List invalidDataList = new ArrayList();
                int errorDataSize = 0;
                for (Object originExcelData : originExcelDataList) {
                    if (originExcelData == null) {
                        continue;
                    }
                    Map<String, List<String>> checkResultMap = excelDataHandler.checkData(originExcelData, this);
                    if (checkResultMap.size() > 0) {
                        hasErrorData = true;
                        importResultDTO.setHasErrorData(hasErrorData);
                        errorDataSize++;
                        generateErrorExcelRow(originExcelData, errorDataSize, columnConfigMap, checkResultMap);
                        invalidDataList.add(originExcelData);
                    } else {
                        validDataList.add(excelDataHandler.fillExtraData(originExcelData));
                    }
                }
                importResultDTO.getSheetResultList().add(ImportResultDTO.SheetResult.builder().excelSheet(excelSheet)
                        .validDataList(validDataList).invalidDataList(invalidDataList).build());
                excelDataHandler.validDataList(validDataList);
                excelDataHandler.invalidDataList(invalidDataList);
            } else {
                importResultDTO.getSheetResultList()
                        .add(ImportResultDTO.SheetResult.builder().excelSheet(excelSheet).build());
            }
        }
    }

    /**
     * 生成错误Excel文件的表头。不管有没有错误数据，都会生成表头，方便编辑后再次导入。
     */
    private void generateErrorExcelHeader() {
        ExcelWriter writer = this.errorExcelExporter.getWriter();
        writer.setSheet(sheetConfig.getIndex());
        writer.renameSheet(sheetConfig.getIndex(), sheetConfig.getName());
        this.errorExcelExporter.fillHeader(columnConfigList);
    }

    /**
     * 生成包含错误信息的Excel行数据。
     *
     * @param originExcelData 原始Excel数据对象
     * @param errorDataSize 错误数据的数量
     * @param columnConfigMap 列配置映射表，用于根据列名获取列索引
     * @param checkResultMap 检查结果映射表，包含错误信息的列名和对应的错误提示列表
     */
    private void generateErrorExcelRow(Object originExcelData, int errorDataSize, Map<String, Integer> columnConfigMap,
        Map<String, List<String>> checkResultMap) {
        ExcelWriter writer = this.errorExcelExporter.getWriter();
        this.errorExcelExporter.fillDropdownRow(columnConfigList, errorDataSize - headerRowNumber);
        this.errorExcelExporter.fillContent(columnConfigList, originExcelData, errorDataSize - headerRowNumber);
        // 根据checkResultMap的key定位列下标，value生成错误提示
        int finalErrorDataSize = errorDataSize;
        checkResultMap.forEach((key, value) -> {
            Integer columnIndex = columnConfigMap.get(key);
            ExcelKit.setCellComment(writer, finalErrorDataSize, columnIndex, value.toString());
            // 给校验异常的单元格设置背景色为醒目红色
            writer.setStyle(ExcelKit.getCheckFailRedStyle(writer), columnIndex, finalErrorDataSize);
        });
    }

    /**
     * 获取Excel中的原始数据。
     *
     * @param sheetDefinition 表格定义类
     * @return Excel中的原始数据列表
     */
    private List getOriginExcelData(Class sheetDefinition) {
        reader.setSheet(sheetConfig.getIndex());
        checkHeader(reader, columnConfigList);
        Map<String, String> headerAlias = new HashMap<>(16);
        for (ExcelConfig.ColumnConfig columnConfig : columnConfigList) {
            headerAlias.put(columnConfig.getHeader(), columnConfig.getFieldName());
        }
        reader.setHeaderAlias(headerAlias);
        // excel中的原始数据
        List<Map<String, Object>> data =
            reader.read(HEADER_INDEX, HEADER_ROW_NUMBER, reader.getSheet().getLastRowNum());
        List result = new ArrayList();
        if (CollectionUtil.isNotEmpty(data)) {
            for (Map<String, Object> map : data) {
                result.add(JSON.parseObject(JSON.toJSONString(map), sheetDefinition));
            }
        }
        return result;
    }

    /**
     * 检查Excel表格的表头是否符合预期。
     *
     * @param reader ExcelReader对象，用于读取Excel表格数据
     * @param columnConfigList 列配置列表，包含每列的索引和预期的表头名称
     * @throws RuntimeException 如果表头不符合预期，则抛出运行时异常
     */
    private void checkHeader(ExcelReader reader, List<ExcelConfig.ColumnConfig> columnConfigList) {
        List<Object> headerList = reader.readRow(HEADER_INDEX);
        Map<Integer, String> headerConfigMap = columnConfigList.stream()
            .collect(Collectors.toMap(ExcelConfig.ColumnConfig::getIndex, ExcelConfig.ColumnConfig::getHeader));
        for (int i = 0; i < headerList.size(); i++) {
            if (!ObjectUtil.equals(headerConfigMap.get(i), headerList.get(i))) {
                throw new RuntimeException(reader.getSheet().getSheetName() + "表头【" + headerList.get(i) + "】应该是【"
                    + headerConfigMap.get(i) + "】");
            }
        }
    }
}
