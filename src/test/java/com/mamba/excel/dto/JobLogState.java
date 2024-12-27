package com.mamba.excel.dto;

import com.mamba.serializer.EnumDefinition;

/**
 * 定时任务日志状态
 *
 * @author wangjie
 */
public enum JobLogState implements EnumDefinition<String> {

    /** 运行 */
    Running("1", "运行中"),
    /** 成功 */
    Success("2", "成功"),
    /** 失败 */
    Failed("3", "失败"),
    /** 数据异常 */
    CompleteWithError("4", "数据异常");

    private final String code;
    private final String comment;

    @Override
    public String getCode() {
        return code;
    }

    @Override
    public String getComment() {
        return comment;
    }

    JobLogState(final String code, final String comment) {
        this.code = code;
        this.comment = comment;
    }
}
