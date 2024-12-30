package com.mamba.excel;

import cn.hutool.core.collection.ListUtil;
import cn.hutool.core.lang.Pair;
import com.mamba.excel.dto.JobLogState;
import com.mamba.excel.dto.PersonDTO;
import com.mamba.excel.dto.PositionDTO;
import com.mamba.excel.kit.ExcelSheetData;
import org.junit.Test;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

/**
 * @author qiangt
 * @version 1.0
 * @date 2024/12/21 19:37
 * @description: 本地磁盘导入导出测试类
 */
public class FoxExcelDiskTest {

    @Test
    public void testWriteAndReadSuccess() {
        String filePath = "D:\\test.xlsx";
        FoxExcel.write(filePath, Pair.of(PersonDTO.class, getPersonList()),
            Pair.of(PositionDTO.class, getPositionList()));
        String errorExcelPath = "D:\\error.xlsx";
        boolean success = FoxExcel.read(filePath, errorExcelPath, PersonDTO.class, PositionDTO.class);
        if (!success) {
            System.out.println("导入excel存在异常数据，详看" + errorExcelPath);
        } else {
            System.out.println("导入excel成功");
        }
    }
    @Test
    public void testWriteAndReadSuccessList() {
        ExcelSheetData<PersonDTO> data1 = new ExcelSheetData<PersonDTO>().setData(getPersonList()).setSheetDefinition(PersonDTO.class);
        ExcelSheetData<PositionDTO> data2 = new ExcelSheetData<PositionDTO>().setData(getPositionList()).setSheetDefinition(PositionDTO.class);
        List<ExcelSheetData> excelSheetDataList = ListUtil.of(data1, data2);
        String filePath = "D:\\test.xlsx";
        FoxExcel.write(filePath, excelSheetDataList);
        String errorExcelPath = "D:\\error.xlsx";
        boolean success = FoxExcel.read(filePath, errorExcelPath, PersonDTO.class, PositionDTO.class);
        if (!success) {
            System.out.println("导入excel存在异常数据，详看" + errorExcelPath);
        } else {
            System.out.println("导入excel成功");
        }
    }

    @Test
    public void testWriteAndReadFailed() {
        String filePath = "D:\\test.xlsx";
        FoxExcel.write(filePath, Pair.of(PersonDTO.class, getErrorPersonList()),
            Pair.of(PositionDTO.class, getErrorPositionList()));
        String errorExcelPath = "D:\\error.xlsx";
        boolean success = FoxExcel.read(filePath, errorExcelPath, PersonDTO.class, PositionDTO.class);
        if (!success) {
            System.out.println("导入excel存在异常数据，详看" + errorExcelPath);
        } else {
            System.out.println("导入excel成功");
        }
    }

    public static List<PersonDTO> getPersonList() {
        PersonDTO person1 = new PersonDTO();
        person1.setName("张三");
        person1.setAge(18);
        person1.setAddress("北京市海淀区颐和园路5号");
        person1.setStaffCode("001");
        person1.setType("A");
        person1.setBirthday(new Date());
        person1.setJoinDay(new Date());
        person1.setCreateTime(LocalDate.of(2024, 12, 21));
        person1.setUpdateTime(LocalDateTime.of(2024, 12, 21, 12, 30));
        person1.setJobLogState(JobLogState.Running);

        PersonDTO person2 = new PersonDTO();
        person2.setName("李四");
        person2.setAge(19);
        person2.setAddress("上海市杨浦区邯郸路220号");
        person2.setStaffCode("002");
        person2.setBirthday(new Date());
        person2.setJoinDay(new Date());
        person2.setCreateTime(LocalDate.of(2024, 12, 21));
        person2.setUpdateTime(LocalDateTime.of(2024, 12, 21, 12, 30));
        person2.setJobLogState(JobLogState.Running);

        return ListUtil.of(person1, person2);
    }

    public static List<PersonDTO> getErrorPersonList() {
        PersonDTO person1 = new PersonDTO();
        person1.setName("张三");
        person1.setAge(188);
        person1.setAddress("北京市海淀区颐和园路5号");
        person1.setStaffCode("001");
        person1.setType("A");
        person1.setBirthday(new Date());
        person1.setJoinDay(new Date());
        person1.setCreateTime(LocalDate.of(2024, 12, 21));
        person1.setUpdateTime(LocalDateTime.of(2024, 12, 21, 12, 30));
        person1.setJobLogState(JobLogState.Running);

        PersonDTO person2 = new PersonDTO();
        person2.setName("李四");
        person2.setAge(19);
        person2.setAddress("上海市杨浦区邯郸路220号");
        person2.setStaffCode("002");
        person2.setBirthday(new Date());
        person2.setJoinDay(new Date());
        person2.setCreateTime(LocalDate.of(2024, 12, 21));
        person2.setUpdateTime(LocalDateTime.of(2024, 12, 21, 12, 30));
        person2.setJobLogState(JobLogState.Running);

        return ListUtil.of(person1, person2);
    }

    public static List<PositionDTO> getPositionList() {
        PositionDTO position1 = new PositionDTO();
        position1.setName("开发工程师");
        position1.setStaffCode("001");
        PositionDTO position2 = new PositionDTO();
        position2.setName("测试工程师");
        position2.setStaffCode("002");
        return ListUtil.of(position1, position2);
    }

    public static List<PositionDTO> getErrorPositionList() {
        PositionDTO position1 = new PositionDTO();
        position1.setName("开发工程师");
        position1.setStaffCode("");
        PositionDTO position2 = new PositionDTO();
        position2.setName("测试工程师");
        position2.setStaffCode("002");
        return ListUtil.of(position1, position2);
    }
}
