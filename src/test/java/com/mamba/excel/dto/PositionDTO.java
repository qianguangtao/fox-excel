package com.mamba.excel.dto;

import com.mamba.excel.annotation.ExcelColumn;
import com.mamba.excel.annotation.ExcelSheet;
import lombok.Data;

import javax.validation.constraints.NotBlank;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 17:47
 * @description: Position excel DTO
 */
@Data
@ExcelSheet(value = "人员职务信息", index = 1)
public class PositionDTO {

    @NotBlank(message = "工号不能为空")
    @ExcelColumn(value = "工号", index = 0, note = "工号备注")
    private String staffCode;
    @ExcelColumn(value = "职务名称", index = 1, note = "职务备注")
    private String name;
}
