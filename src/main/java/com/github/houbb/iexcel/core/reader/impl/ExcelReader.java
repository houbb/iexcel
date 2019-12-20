package com.github.houbb.iexcel.core.reader.impl;

import com.github.houbb.heaven.support.instance.impl.Instances;
import com.github.houbb.heaven.util.guava.Guavas;
import com.github.houbb.heaven.util.lang.StringUtil;
import com.github.houbb.heaven.util.lang.reflect.ClassUtil;
import com.github.houbb.iexcel.annotation.ExcelField;
import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.support.cache.InnerReaderCache;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * excel reader 的基础实现
 * 注意：这个类在阅读的量较大时会存在问题。
 *
 * @author binbin.hou
 * date 2018/11/15 20:00
 * @since 0.0.1
 */
public class ExcelReader<T> implements IExcelReader<T> {

    /**
     * 当前 sheet 的信息
     */
    private final Sheet sheet;
    /**
     * 包含表头
     * 默认第一行是表头(暂时不考虑可以指定哪一行是表头的情况，没有太大实用价值)
     * 如果没有表头可以设置为 false 说明
     */
    private boolean containsHead = true;

    public ExcelReader(File excelFile) {
        this(excelFile, null);
    }

    /**
     * 获取 excel 读取实现
     *
     * @param excelFile  excel 文件信息
     * @param sheetIndex 从0开始
     */
    public ExcelReader(File excelFile, int sheetIndex) {
        Workbook workbook = initWorkbook(excelFile);
        this.sheet = workbook.getSheetAt(sheetIndex);
    }

    public ExcelReader(File excelFile, String sheetName) {
        Workbook workbook = initWorkbook(excelFile);
        if (StringUtil.isBlank(sheetName)) {
            sheetName = ExcelConst.DEFAULT_SHEET_NAME;
        }
        this.sheet = workbook.getSheet(sheetName);
    }

    /**
     * 设置是否包含表头
     *
     * @param containsHead 是否包含表头
     * @return this
     */
    public ExcelReader containsHead(final boolean containsHead) {
        this.containsHead = containsHead;
        return this;
    }

    /**
     * 读取所有的信息
     * 1. 这里的第一行怎么处理 表头的信息？还是按照和导出保持一致？
     *
     * @param tClass 类
     * @return 列表
     */
    @Override
    public List<T> readAll(Class<T> tClass) {
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        return read(tClass, firstRowNum, lastRowNum);
    }

    @Override
    public List<T> read(Class<T> tClass, int startIndex, int endIndex) {
        List<T> resultList = new ArrayList<>();

        int firstRowNum = sheet.getFirstRowNum();
        try {
            //TODO 这里暂时不管，固定是要处理表头的
            // 如果不处理表头，那么 cellFieldMap 固定为 field.name+field 即可。
            // 如果没有表头 则默认和字段含有
            Row firstLineRow = sheet.getRow(firstRowNum);
            Map<Integer, Field> cellFieldMap = getCellFieldMap(tClass, firstLineRow);

            if (containsHead
                && firstRowNum == startIndex) {
                //跳过成为表头处理的那一行
                startIndex++;
            }
            for (int index = startIndex; index <= endIndex; index++) {
                Row row = sheet.getRow(index);
                T instance = tClass.newInstance();

                for(Map.Entry<Integer, Field> entry : cellFieldMap.entrySet()) {
                    final Integer cIndex = entry.getKey();
                    final Field field = entry.getValue();
                    final Class fieldType = field.getType();
                    // 根据字段处理字段信息
                    Cell cell = row.getCell(cIndex);
                    Object cellValue = "";
                    if (null != cell) {
                        cellValue = InnerExcelUtil.getCellValue(cell, cell.getCellTypeEnum(), fieldType);
                    }
                    field.set(instance, cellValue);
                }

                resultList.add(instance);
            }
            return resultList;
        } catch (InstantiationException | IllegalAccessException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    /**
     * 获取 workbook 信息
     *
     * @return workbook
     */
    private Workbook initWorkbook(final File excelFile) {
        try (final InputStream inputStream = new FileInputStream(excelFile)) {
            final String fileName = excelFile.getName();
            if (fileName.endsWith(ExcelTypeEnum.XLS.getValue())) {
                return new HSSFWorkbook(inputStream);
            }
            if (fileName.endsWith(ExcelTypeEnum.XLSX.getValue())) {
                return new XSSFWorkbook(inputStream);
            }
            throw new ExcelRuntimeException("不支持的 excel 文件类型!");
        } catch (IOException e) {
            throw new ExcelRuntimeException("Workbook 初始化异常");
        }
    }

    /**
     * 获取 excel 列信息和 Field 之间的映射关系
     * 1. 如果有表头，则严格要求 bean 中的表头名称，在 excel 中都有。
     *
     * @param tClass       类信息
     * @param firstLineRow 第一行的信息
     * @return 映射 map  Integer 对应的是 cell 的实际列数， Field 是指对应的 bean 字段信息
     */
    private Map<Integer, Field> getCellFieldMap(final Class<?> tClass,
                                                final Row firstLineRow) {
        Map<Integer, Field> cellFieldMap = new HashMap<>();
        final int firstCellNum = firstLineRow.getFirstCellNum();
        final int lastCellNum = firstLineRow.getLastCellNum();

        Map<String, Field> readRequireFieldMap = Instances.singleton(InnerReaderCache.class)
                .get(tClass);

        if (containsHead) {
            // 包含表头的逻辑出路
            // 如果有表头，则严格要求 bean 中的表头名称，在 excel 中都有。
            // 为什么要有这个强制的约束？如果 bean 中的字段,excel中却没有。那么这个 readRequire 应该被声明为 false;
            Map<String, Integer> cellHeadNameIndexMap = cellHeadNameIndexMap(firstLineRow);
            for(Map.Entry<String, Field> entry : readRequireFieldMap.entrySet()) {
                final String headName = entry.getKey();
                final Field field = entry.getValue();
                Integer index = cellHeadNameIndexMap.get(headName);
                if (index != null) {
                    cellFieldMap.put(index, field);
                }
            }
            // 判断是否所有的字段都有对应的 cell
            if (cellFieldMap.size() < readRequireFieldMap.size()) {
                throw new ExcelRuntimeException("指定的表头字段信息和 excel 不符合");
            }
            return cellFieldMap;
        }

        // 这个列数需要是 实际 excel 的列和实际 field 的列数的最小值。
        Integer cellIndex = firstCellNum;
        Collection<Field> fieldList = readRequireFieldMap.values();
        for (Field field : fieldList) {
            cellFieldMap.put(cellIndex, field);
            cellIndex++;
            // 已经达到了 excel 的最后一列
            if (cellIndex > lastCellNum) {
                break;
            }
        }

        return cellFieldMap;
    }

    /**
     * 表头名称和 index 集合
     *
     * @param headRow 表头
     * @return map
     * @since 0.0.7 显式指定 map size，避免扩容消耗。
     */
    private Map<String, Integer> cellHeadNameIndexMap(final Row headRow) {
        final int firstCellNum = headRow.getFirstCellNum();
        final int lastCellNum = headRow.getLastCellNum();
        final int size = lastCellNum-firstCellNum;

        Map<String, Integer> map = Guavas.newHashMap(size);

        for (int i = firstCellNum; i < lastCellNum; i++) {
            map.put(headRow.getCell(i).getStringCellValue(), i);
        }

        return map;
    }

//    /**
//     * 获取需要读取的字段 map
//     *
//     * @param tClass 当前类信息
//     * @return map
//     */
//    @Deprecated
//    private Map<String, Field> readRequireFieldMap(final Class<?> tClass) {
//        Map<String, Field> map = new HashMap<>();
//        List<Field> fieldList = ClassUtil.getAllFieldList(tClass);
//        for (Field field : fieldList) {
//            if (field.isAnnotationPresent(ExcelField.class)) {
//                ExcelField excelField = field.getAnnotation(ExcelField.class);
//                boolean readRequire = excelField.readRequire();
//                if (readRequire) {
//                    final String headName = InnerExcelUtil.getFieldHeadName(excelField, field);
//                    map.put(headName, field);
//                }
//            } else {
//                //@since0.0.4 默认使用 fieldName
//                final String headName = field.getName();
//                map.put(headName, field);
//            }
//        }
//        return map;
//    }

}
