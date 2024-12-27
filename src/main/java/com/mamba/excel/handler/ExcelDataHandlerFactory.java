package com.mamba.excel.handler;

import cn.hutool.core.util.ObjectUtil;
import com.mamba.utils.ReflectUtil;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 11:49
 * @description: excel导入数据处理器工厂类
 */
public class ExcelDataHandlerFactory {

    static {
        Map<String, AbstractExcelDataHandler> map = new HashMap<>(16);
        // 取类包名.分割后的第一个作为待扫描包
        String packageName = ExcelDataHandlerFactory.class.getName().split("\\.")[0];
        Set<Class<? extends AbstractExcelDataHandler>> classSet =
            ReflectUtil.scanClassBySuper(packageName, AbstractExcelDataHandler.class);
        classSet.forEach(clazz -> {
            AbstractExcelDataHandler excelDataHandler = null;
            try {
                excelDataHandler = clazz.newInstance();
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            if (ObjectUtil.isNotNull(excelDataHandler)) {
                map.put(excelDataHandler.getDataClazz(), excelDataHandler);
            }
        });
        beanMap = map;
    }
    static Map<String, AbstractExcelDataHandler> beanMap;

    /**
     * 根据类获取对应的Excel数据处理器
     *
     * @param clazz 要处理的类的Class对象
     * @return 返回对应的Excel数据处理器
     * @throws RuntimeException 如果未找到对应的列数据配置类，则抛出此异常
     */
    public static AbstractExcelDataHandler getExcelDataHandler(Class clazz) {
        AbstractExcelDataHandler excelDataHandler = beanMap.get(clazz.getName());
        if (ObjectUtil.isNull(excelDataHandler)) {
            throw new RuntimeException("未找到对应的列数据配置类");
        }
        return excelDataHandler;
    }
}
