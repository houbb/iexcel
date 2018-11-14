package com.github.houbb.iexcel.support.style;

import org.apache.poi.ss.usermodel.*;

/**
 * 方法直接来自 hutool
 * @author binbin.hou
 * @date 2018/11/14 17:29
 */
public class StyleSet {

    /** 工作簿引用 */
    private Workbook workbook;
    /** 标题样式 */
    protected CellStyle headCellStyle;
    /** 默认样式 */
    protected CellStyle cellStyle;
    /** 默认数字样式 */
    protected CellStyle cellStyleForNumber;
    /** 默认日期样式 */
    protected CellStyle cellStyleForDate;

    /**
     * 构造
     *
     * @param workbook 工作簿
     */
    public StyleSet(Workbook workbook) {
        this.workbook = workbook;
        this.headCellStyle = StyleUtil.createHeadCellStyle(workbook);
        this.cellStyle = StyleUtil.createDefaultCellStyle(workbook);

        // 默认日期格式
        this.cellStyleForDate = StyleUtil.cloneCellStyle(workbook, this.cellStyle);
        // 22表示：m/d/yy h:mm
        this.cellStyleForDate.setDataFormat((short) 22);

        // 默认数字格式
        cellStyleForNumber = StyleUtil.cloneCellStyle(workbook, this.cellStyle);
        // 2表示：0.00
        cellStyleForNumber.setDataFormat((short) 2);
    }

    /**
     * 获取头部样式，获取后可以定义整体头部样式
     *
     * @return 头部样式
     */
    public CellStyle getHeadCellStyle() {
        return headCellStyle;
    }

    /**
     * 获取常规单元格样式，获取后可以定义整体头部样式
     *
     * @return 常规单元格样式
     */
    public CellStyle getCellStyle() {
        return cellStyle;
    }

    /**
     * 获取数字（带小数点）单元格样式，获取后可以定义整体头部样式
     *
     * @return 数字（带小数点）单元格样式
     */
    public CellStyle getCellStyleForNumber() {
        return cellStyleForNumber;
    }

    /**
     * 获取日期单元格样式，获取后可以定义整体头部样式
     *
     * @return 日期单元格样式
     */
    public CellStyle getCellStyleForDate() {
        return cellStyleForDate;
    }

    /**
     * 定义所有单元格的边框类型
     *
     * @param borderSize 边框粗细{@link BorderStyle}枚举
     * @param colorIndex 颜色的short值
     * @return this
     * @since 4.0.0
     */
    public StyleSet setBorder(BorderStyle borderSize, IndexedColors colorIndex) {
        StyleUtil.setBorder(this.headCellStyle, borderSize, colorIndex);
        StyleUtil.setBorder(this.cellStyle, borderSize, colorIndex);
        StyleUtil.setBorder(this.cellStyleForNumber, borderSize, colorIndex);
        StyleUtil.setBorder(this.cellStyleForDate, borderSize, colorIndex);
        return this;
    }

    /**
     * 设置cell文本对齐样式
     *
     * @param halign 横向位置
     * @param valign 纵向位置
     * @return this
     * @since 4.0.0
     */
    public StyleSet setAlign(HorizontalAlignment halign, VerticalAlignment valign) {
        StyleUtil.setAlign(this.headCellStyle, halign, valign);
        StyleUtil.setAlign(this.cellStyle, halign, valign);
        StyleUtil.setAlign(this.cellStyleForNumber, halign, valign);
        StyleUtil.setAlign(this.cellStyleForDate, halign, valign);
        return this;
    }

    /**
     * 设置单元格背景样式
     *
     * @param backgroundColor 背景色
     * @param withHeadCell 是否也定义头部样式
     * @return this
     * @since 4.0.0
     */
    public StyleSet setBackgroundColor(IndexedColors backgroundColor, boolean withHeadCell) {
        if (withHeadCell) {
            StyleUtil.setColor(this.headCellStyle, backgroundColor, FillPatternType.SOLID_FOREGROUND);
        }
        StyleUtil.setColor(this.cellStyle, backgroundColor, FillPatternType.SOLID_FOREGROUND);
        StyleUtil.setColor(this.cellStyleForNumber, backgroundColor, FillPatternType.SOLID_FOREGROUND);
        StyleUtil.setColor(this.cellStyleForDate, backgroundColor, FillPatternType.SOLID_FOREGROUND);
        return this;
    }

    /**
     * 设置全局字体
     *
     * @param color 字体颜色
     * @param fontSize 字体大小，-1表示默认大小
     * @param fontName 字体名，null表示默认字体
     * @param ignoreHead 是否跳过头部样式
     * @return this
     */
    public StyleSet setFont(short color, short fontSize, String fontName, boolean ignoreHead) {
        final Font font = StyleUtil.createFont(this.workbook, color, fontSize, fontName);
        return setFont(font, ignoreHead);
    }

    /**
     * 设置全局字体
     *
     * @param font 字体，可以通过{@link StyleUtil#createFont(Workbook, short, short, String)}创建
     * @param ignoreHead 是否跳过头部样式
     * @return this
     * @since 4.1.0
     */
    public StyleSet setFont(Font font, boolean ignoreHead) {
        if(!ignoreHead) {
            this.headCellStyle.setFont(font);
        }
        this.cellStyle.setFont(font);
        this.cellStyleForNumber.setFont(font);
        this.cellStyleForDate.setFont(font);
        return this;
    }

}
