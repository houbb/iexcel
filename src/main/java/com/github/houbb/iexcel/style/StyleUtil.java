package com.github.houbb.iexcel.style;

import org.apache.poi.ss.usermodel.*;

/**
 * 方法直接来自 hutool
 * @author binbin.hou
 * date 2018/11/14 17:30
 * @since 0.0.1
 */
public final class StyleUtil {

    private StyleUtil(){}

    /**
     * 克隆新的{@link CellStyle}
     *
     * @param workbook 工作簿
     * @param cellStyle 被复制的样式
     * @return {@link CellStyle}
     */
    public static CellStyle cloneCellStyle(Workbook workbook, CellStyle cellStyle) {
        final CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.cloneStyleFrom(cellStyle);
        return newCellStyle;
    }

    /**
     * 设置cell文本对齐样式
     *
     * @param cellStyle {@link CellStyle}
     * @param halign 横向位置
     * @param valign 纵向位置
     * @return {@link CellStyle}
     */
    public static CellStyle setAlign(CellStyle cellStyle, HorizontalAlignment halign, VerticalAlignment valign) {
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        return cellStyle;
    }

    /**
     * 设置cell的四个边框粗细和颜色
     *
     * @param cellStyle {@link CellStyle}
     * @param borderSize 边框粗细{@link BorderStyle}枚举
     * @param colorIndex 颜色的short值
     * @return {@link CellStyle}
     */
    public static CellStyle setBorder(CellStyle cellStyle, BorderStyle borderSize, IndexedColors colorIndex) {
        cellStyle.setBorderBottom(borderSize);
        cellStyle.setBottomBorderColor(colorIndex.index);

        cellStyle.setBorderLeft(borderSize);
        cellStyle.setLeftBorderColor(colorIndex.index);

        cellStyle.setBorderRight(borderSize);
        cellStyle.setRightBorderColor(colorIndex.index);

        cellStyle.setBorderTop(borderSize);
        cellStyle.setTopBorderColor(colorIndex.index);

        return cellStyle;
    }

    /**
     * 给cell设置颜色
     *
     * @param cellStyle {@link CellStyle}
     * @param color 背景颜色
     * @param fillPattern 填充方式 {@link FillPatternType}枚举
     * @return {@link CellStyle}
     */
    public static CellStyle setColor(CellStyle cellStyle, IndexedColors color, FillPatternType fillPattern) {
        return setColor(cellStyle, color.index, fillPattern);
    }

    /**
     * 给cell设置颜色
     *
     * @param cellStyle {@link CellStyle}
     * @param color 背景颜色
     * @param fillPattern 填充方式 {@link FillPatternType}枚举
     * @return {@link CellStyle}
     */
    public static CellStyle setColor(CellStyle cellStyle, short color, FillPatternType fillPattern) {
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(fillPattern);
        return cellStyle;
    }

    /**
     * 创建默认普通单元格样式
     *
     * <pre>
     * 1. 文字上下左右居中
     * 2. 细边框，黑色
     * </pre>
     *
     * @param workbook {@link Workbook} 工作簿
     * @return {@link CellStyle}
     */
    public static CellStyle createDefaultCellStyle(Workbook workbook) {
        final CellStyle cellStyle = workbook.createCellStyle();
        setAlign(cellStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(cellStyle, BorderStyle.THIN, IndexedColors.BLACK);
        return cellStyle;
    }

    /**
     * 创建默认头部样式
     *
     * @param workbook {@link Workbook} 工作簿
     * @return {@link CellStyle}
     */
    public static CellStyle createHeadCellStyle(Workbook workbook) {
        final CellStyle cellStyle = workbook.createCellStyle();
        setAlign(cellStyle, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        setBorder(cellStyle, BorderStyle.THIN, IndexedColors.BLACK);
        setColor(cellStyle, IndexedColors.GREY_25_PERCENT, FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }
}
