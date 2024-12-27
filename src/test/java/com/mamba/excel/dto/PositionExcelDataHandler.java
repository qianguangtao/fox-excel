package com.mamba.excel.dto;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.collection.ListUtil;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.mamba.excel.ExcelImporter;
import com.mamba.excel.handler.AbstractExcelDataHandler;
import lombok.NoArgsConstructor;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 10:38
 * @description: PositionDTO 导入处理器
 */
@NoArgsConstructor
public class PositionExcelDataHandler extends AbstractExcelDataHandler<PositionDTO> {

    @Override
    public String getDataClazz() {
        return PositionDTO.class.getName();
    }

    @Override
    public Map<String, List<String>> checkData(PositionDTO positionDTO, ExcelImporter importer) {
        Map<String, List<String>> resultMap = new HashMap<>(16);
        Map<String, List<String>> checkMap = this.validateData(positionDTO);
        if (CollectionUtil.isNotEmpty(checkMap)) {
            resultMap.putAll(checkMap);
        }
        List<PositionDTO> allDataList = importer.getAllDataMap().get(PositionDTO.class.getName());
        long count =
                allDataList.stream().filter(item -> item.getStaffCode().equals(positionDTO.getStaffCode())).count();
        if (count > 1) {
            String staffCode = "staffCode";
            String errorMessage = "员工编号" + positionDTO.getStaffCode() + "重复";
            if (resultMap.containsKey(staffCode)) {
                resultMap.get(staffCode).add(errorMessage);
            } else {
                resultMap.put(staffCode, ListUtil.of(errorMessage));
            }
        }
        return resultMap;
    }

    @Override
    public void validDataList(List<PositionDTO> validDataList) {
        // 数据入库
        System.out.println("数据入库");
        System.out.println(JSON.toJSONString(validDataList, SerializerFeature.PrettyFormat));
    }

    @Override
    public void invalidDataList(List<PositionDTO> invalidDataList) {
        System.out.println("异常数据");
        System.out.println(JSON.toJSONString(invalidDataList, SerializerFeature.PrettyFormat));
    }
}
