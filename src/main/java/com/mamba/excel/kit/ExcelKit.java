package com.mamba.excel.kit;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.util.List;
import java.util.Optional;

/**
 * @author 00351634
 * @version 1.0
 * @date 2024/12/20 19:04
 * @description: excel工具类
 */
public class ExcelKit {
    /**
     * 设置单元格备注
     *
     * @param writer
     * @param row
     * @param col
     * @param note
     */
    public static void setCellComment(ExcelWriter writer, int row, int col, String note) {
        if (StrUtil.isBlank(note)) {
            return;
        }
        // 创建单元格备注的位置
        Sheet sheet = writer.getSheet();
        Cell cell = writer.getCell(col, row);
        // 判断单元格如果已经有备注则删除已有的备注
        if (cell.getCellComment() != null) {
            cell.removeCellComment();
        }

        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper factory = writer.getWorkbook().getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        // 备注的起始行
        anchor.setRow1(row);
        // 备注的结束行
        anchor.setRow2(row);
        // 备注的起始列
        anchor.setCol1(col);
        // 备注的结束列
        anchor.setCol2(col);
        // 创建备注
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(new XSSFRichTextString(note));
        comment.setAuthor("系统提示");
        // 将备注添加到单元格
        cell.setCellComment(comment);
    }

    /**
     * 校验异常的单元格样式-红色背景
     *
     * @param writer
     * @return
     */
    public static CellStyle getCheckFailRedStyle(ExcelWriter writer) {
        CellStyle style = writer.createCellStyle();
        Font font = writer.createFont();
        // 设置字体名称 宋体 / 微软雅黑 /等
        font.setFontName("Consolas");
        // 是否加粗
        font.setBold(false);
        // 设置是否斜体
        font.setItalic(false);
        // 设置字体高度
        // font.setFontHeight((short) fontHeight);
        // 设置字体大小 以磅为单位
        font.setFontHeightInPoints((short)11);
        // 默认字体颜色 (红色 Font.COLOR_RED)
        font.setColor(Font.COLOR_NORMAL);
        // 设置下划线样式
        // font.setUnderline(Font.ANSI_CHARSET);
        // 设定文字删除线
        font.setStrikeout(false);
        style.setFont(font);

        // 是否自动换行
        style.setWrapText(false);
        // 垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 水平居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setLocked(false);
        // 顶部边框
        style.setBorderTop(BorderStyle.THIN);
        // 底部边框
        style.setBorderBottom(BorderStyle.THIN);
        // 左边框
        style.setBorderLeft(BorderStyle.THIN);
        // 右边框
        style.setBorderRight(BorderStyle.THIN);
        // 填充颜色
        // 设置背景色 红色
        style.setFillForegroundColor(IndexedColors.RED.getIndex());
        // 必须设置 否则背景色不生效
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    /**
     * 为ExcelWriter对象中的所有Sheet设置自动列宽。
     *
     * @param writer ExcelWriter对象，用于获取所有的Sheet。
     */
    public static void setAutoSizeColumn(ExcelWriter writer) {
        List<Sheet> sheets = writer.getSheets();
        if (CollectionUtil.isNotEmpty(sheets)) {
            for (Sheet sheet : sheets) {
                setAutoSizeColumn(sheet);
            }
        }
    }

    /**
     * 自适应宽度(中文支持)
     *
     * @param sheet
     */
    public static void setAutoSizeColumn(Sheet sheet) {
        // 获取工作表中的所有行
        for (Row row : sheet) {
            // 遍历行中的所有单元格
            for (Cell cell : row) {
                // 设置单元格的类型为字符串
                // cell.setCellType(CellType.STRING);
                // 自适应列宽
                int columnNum = cell.getColumnIndex();
                int columnWidth = sheet.getColumnWidth(columnNum) / 256;
                if (cell.getCellType() == CellType.STRING) {
                    int length = cell.getStringCellValue().getBytes().length;
                    if (columnWidth < length) {
                        columnWidth = length;
                    }
                }
                sheet.setColumnWidth(columnNum, columnWidth * 256);
            }
        }
    }

    /**
     * 设置下拉列表
     *
     * @param sheet Excel表格对象
     * @param firstRow 下拉列表起始行
     * @param firstCol 下拉列表起始列
     * @param lastRow 下拉列表结束行
     * @param lastCol 下拉列表结束列
     * @param values 下拉列表选项数组
     */
    public static void setDropdownList(Sheet sheet, int firstRow, int firstCol, int lastRow, int lastCol,
                                       String[] values) {
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(values);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation validation = validationHelper.createValidation(constraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }
}
