package com.github.houbb.iexcel.util.excel;

import com.github.houbb.iexcel.support.style.StyleSet;
import com.github.houbb.iexcel.util.StrUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;

import java.util.Calendar;
import java.util.Date;

/**
 * 方法直接来自 hutool
 * Excel表格中单元格工具类
 * @author binbin.hou
 * @date 2018/11/14 14:39
 */
public class CellUtil {

    /**
     * 设置单元格值<br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell 单元格
     * @param value 值
     * @param styleSet 单元格样式集，包括日期等样式
     * @param isHeader 是否为标题单元格
     */
    public static void setCellValue(Cell cell, Object value, StyleSet styleSet, boolean isHeader) {
        final CellStyle headCellStyle = styleSet.getHeadCellStyle();
        final CellStyle cellStyle = styleSet.getCellStyle();
        if(isHeader && null != headCellStyle) {
            cell.setCellStyle(headCellStyle);
        } else if (null != cellStyle) {
            cell.setCellStyle(cellStyle);
        }

        if (null == value) {
            cell.setCellValue(StrUtil.EMPTY);
        } else if (value instanceof Date) {
            if (null != styleSet.getCellStyleForDate()) {
                cell.setCellStyle(styleSet.getCellStyleForDate());
            }
            cell.setCellValue((Date) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        } else if (value instanceof Number) {
            if ((value instanceof Double || value instanceof Float) && null != styleSet && null != styleSet.getCellStyleForNumber()) {
                cell.setCellStyle(styleSet.getCellStyleForNumber());
            }
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            cell.setCellValue(value.toString());
        }
    }

}
