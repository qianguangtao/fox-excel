package com.mamba.excel.dto;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.mamba.excel.ExcelImporter;
import com.mamba.excel.handler.AbstractExcelDataHandler;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/21 10:38
 * @description: PersonDTO 导入处理器
 */
@NoArgsConstructor
public class PersonExcelDataHandler extends AbstractExcelDataHandler<PersonDTO> {

    @Override
    public String getDataClazz() {
        return PersonDTO.class.getName();
    }

    @Override
    public PersonDTO fillExtraData(PersonDTO personDTO) {
        personDTO.setType("1");
        return personDTO;
    }

    @Override
    public Map<String, List<String>> checkData(PersonDTO personDTO, ExcelImporter importer) {
        return this.validateData(personDTO);
    }

    @Override
    public void validDataList(List<PersonDTO> validDataList) {
        // 数据入库
        System.out.println("数据入库");
        System.out.println(JSON.toJSONString(validDataList, SerializerFeature.PrettyFormat));
    }

    @Override
    public void invalidDataList(List<PersonDTO> invalidDataList) {
        System.out.println("异常数据" + invalidDataList);
        System.out.println(JSON.toJSONString(invalidDataList, SerializerFeature.PrettyFormat));
    }
}
