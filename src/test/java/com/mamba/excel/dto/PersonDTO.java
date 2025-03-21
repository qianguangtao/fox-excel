package com.mamba.excel.dto;

import com.alibaba.fastjson.annotation.JSONField;
import com.mamba.excel.annotation.ExcelColumn;
import com.mamba.excel.annotation.ExcelSheet;
import com.mamba.serializer.EnumConverter;
import lombok.Data;
import org.springframework.format.annotation.DateTimeFormat;

import javax.validation.constraints.Max;
import javax.validation.constraints.NotBlank;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:43
 * @description: Person excel DTO
 */
@Data
@ExcelSheet(value = "人员信息", index = 0)
public class PersonDTO {

    @NotBlank(message = "姓名不能为空")
    @ExcelColumn(value = "姓名", index = 0, note = "姓名备注")
    private String name;
    @Max(value = 100, message = "年龄不能超过100")
    @ExcelColumn(value = "年龄", index = 1, note = "年龄备注")
    private Integer age;
    @ExcelColumn(value = "地址", index = 2)
    private String address;
    @ExcelColumn(value = "工号", index = 3, note = "工号备注")
    private String staffCode;
    /** 1: 内部员工，2：外部员工。模拟非导入数据 */
    private String type;

    @ExcelColumn(value = "生日", index = 4)
    private Date birthday;

    @ExcelColumn(value = "入职日期", index = 5)
    @DateTimeFormat(pattern = "yyyy-MM-dd")
    private Date joinDay;

    @ExcelColumn(value = "创建时间", index = 6)
    private LocalDate createTime;

    @ExcelColumn(value = "更新时间", index = 7)
    private LocalDateTime updateTime;

    @ExcelColumn(value = "测试枚举", index = 8, enumDefinition = JobLogState.class)
    @JSONField(serializeUsing = EnumConverter.class, deserializeUsing = EnumConverter.class)
    private JobLogState jobLogState;
}
