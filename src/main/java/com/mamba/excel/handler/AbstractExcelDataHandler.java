package com.mamba.excel.handler;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.ReflectUtil;
import com.mamba.excel.ExcelImporter;
import com.mamba.utils.ValidateUtils;

import javax.validation.ConstraintViolation;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 10:24
 * @description: 抽象excel导入数据处理器
 *
 */
public abstract class AbstractExcelDataHandler<T> {

    /**
     * 获取数据类型的Class全路径。子类需要实现此方法，返回泛型T的Class全路径
     *
     * @return 返回数据类型的Class全路径
     */
    abstract public String getDataClazz();

    /**
     * excel中的数据采集后，填充额外数据
     *
     * @param t
     * @return
     */
    public T fillExtraData(T t) {
        return t;
    }

    /**
     * 校验数据，返回校验不通过的字段和校验不通过的原因
     *
     * @param t
     * @param importer，excel导入器，可获取全部的导入数据，用于全量数据校验
     * @return key 为字段名，value 为校验不通过的原因
     */
    public abstract Map<String, List<String>> checkData(T t, ExcelImporter importer);

    /**
     * 使用javax.validation注解验证指定对象的数据有效性
     *
     * @param t 需要验证的对象
     * @return 包含验证结果和错误信息的Map集合，键为错误字段名称，值为对应的错误信息列表
     */
    protected Map<String, List<String>> validateData(T t) {
        Map<String, List<String>> map = new HashMap<>();
        Field[] fields = ReflectUtil.getFieldsDirectly(t.getClass(), false);
        for (Field field : fields) {
            Set<ConstraintViolation<Object>> violations = ValidateUtils.validateProperty(t, field.getName());
            if (CollectionUtil.isNotEmpty(violations)) {
                map.put(field.getName(),
                    violations.stream().map(ConstraintViolation::getMessage).collect(Collectors.toList()));
            }
        }
        return map;
    }

    /**
     * 获取正常数据
     *
     * @param validDataList
     */
    public abstract void validDataList(List<T> validDataList);

    /**
     * 获取异常数据，默认不处理
     *
     * @param invalidDataList
     */
    public void invalidDataList(List<T> invalidDataList) {}

}
