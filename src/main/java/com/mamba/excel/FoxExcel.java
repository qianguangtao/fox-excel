package com.mamba.excel;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.lang.Pair;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.StrUtil;
import com.mamba.excel.kit.ExcelSheetData;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.function.Supplier;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/24 13:40
 * @description: 导入导出工具类
 */
public class FoxExcel {

    /**
     * 默认导入成功response
     */
    private static final Supplier DEFAULT_SUCCESS = () -> {
        return MapUtil.of("msg", "导入成功");
    };

    /**
     * 从MultipartFile文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到名为"异常-<原始文件名>"的Excel文件中，并返回成功响应。
     *
     * @param file 文件对象，MultipartFile类型
     * @param response HttpServletResponse对象，用于返回导入结果
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     */
    public static void read(MultipartFile file, HttpServletResponse response, Class... sheetDefinition) {
        read(file, response, null, DEFAULT_SUCCESS, sheetDefinition);

    }

    /**
     * 从MultipartFile文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到指定的错误Excel文件中，并返回成功响应。
     *
     * @param file 文件对象，MultipartFile类型
     * @param response HttpServletResponse对象，用于返回导入结果
     * @param errorExcelName 错误Excel文件的名称，如果为null或空字符串，则使用默认名称"异常-<原始文件名>"
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     */
    public static void read(MultipartFile file, HttpServletResponse response, String errorExcelName,
        Class... sheetDefinition) {
        read(file, response, errorExcelName, DEFAULT_SUCCESS, sheetDefinition);
    }

    /**
     * 从MultipartFile文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到指定的错误Excel文件中，并返回由success参数指定的成功响应。
     *
     * @param file 文件对象，MultipartFile类型
     * @param response HttpServletResponse对象，用于返回导入结果
     * @param errorExcelName 错误Excel文件的名称，如果为null或空字符串，则使用默认名称"异常-<原始文件名>"
     * @param success 成功时返回的响应数据提供者
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     */
    public static void read(MultipartFile file, HttpServletResponse response, String errorExcelName, Supplier success,
        Class... sheetDefinition) {
        Assert.notEmpty(sheetDefinition);
        ExcelImporter importer = new ExcelImporter(file);
        importer.importData(Arrays.asList(sheetDefinition), response,
            StrUtil.blankToDefault(errorExcelName, "异常-" + file.getOriginalFilename()), success);
    }

    /**
     * 从指定路径的文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到指定的错误Excel文件中。
     *
     * @param filePath 文件路径，表示要读取数据的Excel文件位置
     * @param errorExcelPath 错误Excel文件的路径，如果导入过程中存在错误数据，则会将这些数据导出到这个文件中
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     * @return 如果导入过程中没有错误数据，则返回true；否则返回false
     * @throws IllegalArgumentException 如果sheetDefinition为空，则抛出此异常
     * @throws IllegalArgumentException 如果文件不存在，则抛出此异常
     */
    public static boolean read(String filePath, String errorExcelPath, Class... sheetDefinition) {
        Assert.notEmpty(sheetDefinition);
        File file = new File(filePath);
        Assert.notNull(file, "文件不存在");
        ExcelImporter importer = new ExcelImporter(filePath);
        importer.importData(Arrays.asList(sheetDefinition), errorExcelPath);
        return !importer.isHasErrorData();
    }

    /**
     * 将数据写入Excel文件并返回HttpServletResponse。
     *
     * @param fileName Excel文件的名称
     * @param excelSheetDataPairs 包含Sheet定义类和数据列表的Pair对象数组，用于指定每个Sheet的数据内容
     */
    public static void write(String fileName, HttpServletResponse response, Pair<Class, List>... excelSheetDataPairs) {
        Assert.notEmpty(excelSheetDataPairs);
        List<ExcelSheetData> excelSheetDataList = getExcelSheetDataList(excelSheetDataPairs);
        ExcelExporter excelExporter = new ExcelExporter();
        excelExporter.exportData(excelSheetDataList, response, fileName);
    }

    /**
     * 将数据写入本地磁盘Excel文件
     *
     * @param filePath Excel文件的路径（本地磁盘路径+文件名）
     * @param excelSheetDataPairs 包含Class和List对的数组，每个对表示一个Excel表格的数据
     * @throws IllegalArgumentException 如果excelSheetDataPairs为空
     */
    public static void write(String filePath, Pair<Class, List>... excelSheetDataPairs) {
        Assert.notEmpty(excelSheetDataPairs);
        List<ExcelSheetData> excelSheetDataList = getExcelSheetDataList(excelSheetDataPairs);
        ExcelExporter excelExporter = new ExcelExporter();
        excelExporter.exportData(excelSheetDataList, filePath);
    }

    /**
     * 将Pair数组转换为ExcelSheetData列表
     *
     * @param excelSheetDataPairs 包含Class和List对的数组，用于生成ExcelSheetData对象
     * @return 包含所有ExcelSheetData对象的列表
     */
    private static List<ExcelSheetData> getExcelSheetDataList(Pair<Class, List>... excelSheetDataPairs) {
        List<ExcelSheetData> excelSheetDataList = new ArrayList();
        for (Pair pair : excelSheetDataPairs) {
            ExcelSheetData excelSheetData = new ExcelSheetData();
            excelSheetData.setSheetDefinition((Class)pair.getKey());
            excelSheetData.setData(
                Optional.ofNullable(pair.getValue()).map(v -> (List)v).orElseGet(CollectionUtil::newArrayList));
            excelSheetDataList.add(excelSheetData);
        }
        return excelSheetDataList;
    }
}
