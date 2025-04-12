package com.mamba.excel.kit;

import com.mamba.excel.annotation.ExcelSheet;
import lombok.Builder;
import lombok.Data;
import org.apache.commons.compress.utils.Lists;

import java.util.ArrayList;
import java.util.List;

/**
 * @author 00351634
 * @version 1.0
 * @date 2025/3/1 20:00
 * @description: 导入结果DTO
 */
@Data
public class ImportResultDTO {
    /** 是否包含错误数据 */
    private boolean hasErrorData;
    /** 导入结果详情（分sheet） */
    private List<SheetResult> sheetResultList;

    public ImportResultDTO() {
        super();
        hasErrorData = false;
        sheetResultList = Lists.newArrayList();
    }

    @Data
    @Builder
    public static class SheetResult {
        /** 导入Sheet的注解定义 */
        private ExcelSheet excelSheet;
        /** 导入成功的数据 */
        private List validDataList = new ArrayList();
        /** 导入失败的数据 */
        private List invalidDataList = new ArrayList();
    }
}
