package com.mamba.excel;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.lang.Pair;
import cn.hutool.core.map.MapUtil;
import cn.hutool.core.util.StrUtil;
import com.mamba.excel.kit.ExcelSheetData;
import com.mamba.excel.kit.ImportResultDTO;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.function.BiFunction;
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
     * 默认导入结果处理函数，不做任何操作
     */
    private static final BiFunction<ImportResultDTO, ExcelExporter, Boolean> DEFAULT_RESULT_FUNCTION = (result, errorExcelExporter) -> {return false;
    };

    /**
     * 从MultipartFile文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到名为"异常-<原始文件名>"的Excel文件中，并返回成功响应。
     *
     * @param file 文件对象，MultipartFile类型
     * @param response HttpServletResponse对象，用于返回导入结果
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     */
    public static void read(MultipartFile file, HttpServletResponse response, Class... sheetDefinition) {
        read(file, response, null, DEFAULT_SUCCESS, DEFAULT_RESULT_FUNCTION, sheetDefinition);

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
        read(file, response, errorExcelName, DEFAULT_SUCCESS, DEFAULT_RESULT_FUNCTION, sheetDefinition);
    }

    /**
     * 从MultipartFile文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到指定的错误Excel文件中，并返回由success参数指定的成功响应。
     *
     * @param file 文件对象，MultipartFile类型
     * @param response HttpServletResponse对象，用于返回导入结果
     * @param errorExcelName 错误Excel文件的名称，如果为null或空字符串，则使用默认名称"异常-<原始文件名>"
     * @param success 成功时返回的响应数据提供者
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     * @param function 用于处理导入结果回调
     */
    public static void read(MultipartFile file, HttpServletResponse response, String errorExcelName, Supplier success,
                            BiFunction<ImportResultDTO, ExcelExporter, Boolean> function,
                            Class... sheetDefinition) {
        Assert.notEmpty(sheetDefinition);
        ExcelImporter importer = new ExcelImporter(file);
        importer.importData(Arrays.asList(sheetDefinition), response,
                StrUtil.blankToDefault(errorExcelName, "异常-" + file.getOriginalFilename()), success, function);
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
        return read(filePath, errorExcelPath, DEFAULT_RESULT_FUNCTION, sheetDefinition);
    }

    /**
     * 从指定路径的文件中读取数据，并根据提供的Sheet定义类导入数据。 如果存在错误数据，则将其导出到指定的错误Excel文件中。
     *
     * @param filePath 文件路径，表示要读取数据的Excel文件位置
     * @param errorExcelPath 错误Excel文件的路径，如果导入过程中存在错误数据，则会将这些数据导出到这个文件中
     * @param sheetDefinition Sheet定义类数组，用于指定每个Sheet的类定义
     * @param function 用于处理导入结果回调
     * @return 如果导入过程中没有错误数据，则返回true；否则返回false
     * @throws IllegalArgumentException 如果sheetDefinition为空，则抛出此异常
     * @throws IllegalArgumentException 如果文件不存在，则抛出此异常
     */
    public static boolean read(String filePath, String errorExcelPath, BiFunction<ImportResultDTO, ExcelExporter, Boolean> function, Class... sheetDefinition) {
        Assert.notEmpty(sheetDefinition);
        File file = new File(filePath);
        Assert.notNull(file, "文件不存在");
        ExcelImporter importer = new ExcelImporter(filePath);
        importer.importData(Arrays.asList(sheetDefinition), errorExcelPath, function);
        return !importer.isHasErrorData();
    }

    /**
     * 将数据写入Excel文件并返回HttpServletResponse。
     *
     * @param fileName Excel文件的名称
     * @param response HttpServletResponse对象，用于将生成的Excel文件作为响应发送给客户端
     * @param excelSheetDataPairs 包含Sheet定义类和数据列表的Pair对象数组，用于指定每个Sheet的数据内容 每个Pair对象包含一个Sheet的定义类和一个包含数据的List。
     *            定义类需要包含注解，用于指定Sheet的名称、列名等信息。
     * @throws IllegalArgumentException 如果excelSheetDataPairs为空，则抛出此异常
     */
    public static void write(String fileName, HttpServletResponse response, Pair<Class, List>... excelSheetDataPairs) {
        Assert.notEmpty(excelSheetDataPairs);
        write(fileName, response, getExcelSheetDataList(excelSheetDataPairs));
    }

    /**
     * 将Excel数据写入到HTTP响应中。
     *
     * @param fileName Excel文件的名称，包括扩展名。
     * @param response HttpServletResponse对象，用于将生成的Excel文件作为HTTP响应发送给客户端。
     * @param excelSheetDataList 包含要写入Excel文件的数据的列表。每个元素代表一个Excel工作表的数据。
     * @throws IllegalArgumentException 如果excelSheetDataList为空，则抛出此异常。
     */
    public static void write(String fileName, HttpServletResponse response, List<ExcelSheetData> excelSheetDataList) {
        Assert.notEmpty(excelSheetDataList);
        ExcelExporter excelExporter = new ExcelExporter();
        excelExporter.exportData(excelSheetDataList, response, fileName);
    }

    /**
     * 将Excel数据写入本地磁盘Excel文件
     *
     * @param filePath Excel文件的路径（本地磁盘路径+文件名）
     * @param excelSheetDataPairs 包含Class和List对的数组，每个对表示一个Excel表格的数据
     * @throws IllegalArgumentException 如果excelSheetDataPairs为空
     */
    public static void write(String filePath, Pair<Class, List>... excelSheetDataPairs) {
        Assert.notEmpty(excelSheetDataPairs);
        write(filePath, getExcelSheetDataList(excelSheetDataPairs));
    }

    /**
     * 将Excel数据写入本地磁盘Excel文件
     *
     * @param filePath Excel文件的路径（本地磁盘路径+文件名）
     * @param excelSheetDataList 包含Sheet定义类和数据列表的Pair对象数组，用于指定每个Sheet的数据内容
     * @throws IllegalArgumentException 如果excelSheetDataList为空，则抛出此异常。
     */
    public static void write(String filePath, List<ExcelSheetData> excelSheetDataList) {
        Assert.notEmpty(excelSheetDataList);
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
            excelSheetDataList.add(new ExcelSheetData().setSheetDefinition((Class)pair.getKey()).setData(
                    Optional.ofNullable(pair.getValue()).map(v -> (List)v).orElseGet(CollectionUtil::newArrayList)));
        }
        return excelSheetDataList;
    }
}
