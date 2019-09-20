package com.github.houbb.iexcel.util.excel;

import com.github.houbb.heaven.constant.CharConst;
import com.github.houbb.heaven.constant.PunctuationConst;
import com.github.houbb.heaven.util.lang.StringUtil;
import com.github.houbb.heaven.util.lang.reflect.ClassUtil;
import com.github.houbb.iexcel.annotation.ExcelField;
import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.style.StyleSet;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 内部 EXCEL 工具类
 *
 * @author binbin.hou
 * date 2018/11/14 20:06
 * @since 0.0.1
 */
public final class InnerExcelUtil {

    /**
     * 私有化构造器
     * @since 0.0.3
     */
    private InnerExcelUtil(){}

    /**
     * 获取读取 cell 的下标和类字段的映射关系
     * @param targetClass 目标类
     * @param headNameList 表头列表
     * @return 映射关系 map
     */
    public static Map<Integer, Field> getReadIndexFieldMap(final Class<?> targetClass,
                                                          final List<Object> headNameList) {
        Map<Integer, Field> readIndexFieldMap = new HashMap<>();
        Map<String, Field> readRequireFieldMap = readRequireFieldMap(targetClass);

        // 强制约定：如果 readRequire 则 excel 必须有与之对应的字段信息
        for(String headName : readRequireFieldMap.keySet()) {
            int index = headNameList.indexOf(headName);
            readIndexFieldMap.put(index, readRequireFieldMap.get(headName));
        }

        if(readIndexFieldMap.size() != readRequireFieldMap.size()) {
            throw new ExcelRuntimeException("excel 的表头信息和 bean 指定的表头信息不一致");
        }

        return readIndexFieldMap;
    }

    /**
     * 获取需要读取的字段 map
     * @param tClass 当前类信息
     * @return map
     */
    private static Map<String, Field> readRequireFieldMap(final Class<?> tClass) {
        Map<String, Field> map = new HashMap<>();
        List<Field> fieldList = ClassUtil.getAllFieldList(tClass);
        for(Field field : fieldList) {
            if(field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                boolean readRequire = excelField.readRequire();
                if(readRequire) {
                    final String headName = getFieldHeadName(excelField, field);
                    map.put(headName, field);
                }
            }
        }
        return map;
    }


    /**
     * 获取当前字段对应的 headName
     * @param excelField excel 子弹信息
     * @param field 字段
     * @return 对应的 headName
     */
    public static String getFieldHeadName(final ExcelField excelField, final Field field) {
        final String fieldName = field.getName();
        String headName = excelField.headName();
        return StringUtil.isNotBlank(headName) ? headName : fieldName;
    }

    /**
     * 校验列的数量
     *
     * @param size 数量
     */
    public static void checkColumnNum(final int size) {
        if (size <= 0
                || size > ExcelConst.MAX_COLUMNS_LIMIT) {
            throw new ExcelRuntimeException("excel 列数不在正确范围内");
        }
    }

    /**
     * 写一行数据
     *
     * @param row      行
     * @param rowData  一行的数据
     * @param styleSet 单元格样式集，包括日期等样式
     * @param isHeader 是否为标题行
     */
    public static void writeRow(Row row, Iterable<?> rowData, StyleSet styleSet, boolean isHeader) {
        int i = 0;
        Cell cell;
        for (Object value : rowData) {
            cell = row.createCell(i);
            InnerExcelUtil.setCellValue(cell, value, styleSet, isHeader);
            i++;
        }
    }

    /**
     * 设置单元格值<br>
     * 根据传入的styleSet自动匹配样式<br>
     * 当为头部样式时默认赋值头部样式，但是头部中如果有数字、日期等类型，将按照数字、日期样式设置
     *
     * @param cell     单元格
     * @param value    值
     * @param styleSet 单元格样式集，包括日期等样式
     * @param isHeader 是否为标题单元格
     */
    private static void setCellValue(Cell cell, Object value, StyleSet styleSet, boolean isHeader) {
        final CellStyle headCellStyle = styleSet.getHeadCellStyle();
        final CellStyle cellStyle = styleSet.getCellStyle();
        if (isHeader && null != headCellStyle) {
            cell.setCellStyle(headCellStyle);
        } else if (null != cellStyle) {
            cell.setCellStyle(cellStyle);
        }

        if (null == value) {
            cell.setCellValue(StringUtil.EMPTY);
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

    /**
     * 获取单元格值<br>
     * 如果单元格值为数字格式，则判断其格式中是否有小数部分，无则返回Long类型，否则返回Double类型
     *
     * @param cell {@link Cell}单元格
     * @param cellType 单元格值类型{@link CellType}枚举，如果为{@code null}默认使用cell的类型
     * @return 值，类型可能为：Date、Double、Boolean、String
     */
    public static Object getCellValue(Cell cell, CellType cellType) {
        if (null == cell) {
            return null;
        }
        if (null == cellType) {
            cellType = cell.getCellTypeEnum();
        }

        Object value;
        switch (cellType) {
            case NUMERIC:
                value = getNumericValue(cell);
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                // 遇到公式时查找公式结果类型
                value = getCellValue(cell, cell.getCachedFormulaResultTypeEnum());
                break;
            case BLANK:
                value = StringUtil.EMPTY;
                break;
            case ERROR:
                final FormulaError error = FormulaError.forInt(cell.getErrorCellValue());
                value = (null == error) ? StringUtil.EMPTY : error.getString();
                break;
            default:
                value = cell.getStringCellValue();
        }

        return value;
    }

    /**
     * 获取数字类型的单元格值
     *
     * @param cell 单元格
     * @return 单元格值，可能为Long、Double、Date
     */
    private static Object getNumericValue(Cell cell) {
        final double value = cell.getNumericCellValue();

        final CellStyle style = cell.getCellStyle();
        if (null == style) {
            return value;
        }

        final short formatIndex = style.getDataFormat();
        // 判断是否为日期
        if (isDateType(cell, formatIndex)) {
            cell.getDateCellValue();
        }

        final String format = style.getDataFormatString();
        // 普通数字
        if (null != format && format.indexOf(CharConst.DOT) < 0) {
            final long longPart = (long) value;
            if (longPart == value) {
                // 对于无小数部分的数字类型，转为Long
                return longPart;
            }
        }
        return value;
    }

    /**
     * 是否为日期格式<br>
     * 判断方式：
     *
     * <pre>
     * 1、指定序号
     * 2、org.apache.poi.ss.usermodel.DateUtil.isADateFormat方法判定
     * </pre>
     *
     * @param formatIndex 格式序号
     * @param cell 列
     * @return 是否为日期格式
     */
    private static boolean isDateType(Cell cell, int formatIndex) {
        // yyyy-MM-dd----- 14
        // yyyy年m月d日---- 31
        // yyyy年m月------- 57
        // m月d日 ---------- 58
        // HH:mm----------- 20
        // h时mm分 -------- 32
        if (formatIndex == 14 || formatIndex == 31 || formatIndex == 57 || formatIndex == 58 || formatIndex == 20 || formatIndex == 32) {
            return true;
        }

        return DateUtil.isCellDateFormatted(cell);
    }
}
