package com.github.houbb.iexcel.core.writer.impl;


import com.github.houbb.heaven.support.instance.impl.Instances;
import com.github.houbb.heaven.util.common.ArgUtil;
import com.github.houbb.heaven.util.lang.BeanUtil;
import com.github.houbb.heaven.util.lang.StringUtil;
import com.github.houbb.heaven.util.lang.reflect.ClassUtil;
import com.github.houbb.heaven.util.util.CollectionUtil;
import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.hutool.annotation.ExcelField;
import com.github.houbb.iexcel.style.StyleSet;
import com.github.houbb.iexcel.support.cache.HeaderAliasCache;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 基础 excel writer 类
 * @author binbin.hou
 * date 2018/11/14 13:56
 */
public abstract class AbstractExcelWriter implements IExcelWriter {

    /**
     * 当前行
     */
    private AtomicInteger currentRow = new AtomicInteger(0);

    /**
     * 是否被关闭
     */
    private volatile boolean isClosed;

    /**
     * 是否已经包含表头
     * 1. 避免多次写入，表头也被写入多次的问题
     * @since 0.0.3
     */
    private volatile boolean containsHeadRow;

    /**
     * 表头别名信息 map
     */
    private Map<String, String> headerAliasMap;

    /**
     * 样式集，定义不同类型数据样式
     */
    private StyleSet styleSet;

    /**
     * 工作簿
     */
    private Workbook workbook;

    /**
     * Excel中对应的Sheet
     */
    private Sheet sheet;

    protected AbstractExcelWriter() {
        this(null);
    }

    protected AbstractExcelWriter(final String sheetName) {
        this.workbook = getWorkbook();
        final String realSheetName = StringUtil.isBlank(sheetName)
                ? ExcelConst.DEFAULT_SHEET_NAME : sheetName;
        this.sheet = workbook.createSheet(realSheetName);
        this.styleSet = new StyleSet(workbook);
    }

    /**
     * 获取 workbook
     * @return workbook
     */
    protected abstract Workbook getWorkbook();

    /**
     * 获取最大行数限制
     * @return 最大行数限制
     */
    protected abstract int getMaxRowNumLimit();

    @Override
    public IExcelWriter write(Collection<?> data) {
        // 快速返回
        if(CollectionUtil.isEmpty(data)) {
            return this;
        }

        // 状态校验
        checkClosedStatus();
        checkRowNum(data.size());

        // 生成 excel 文件
        // 处理第一条数据信息
        Iterator iterator = data.iterator();
        Object firstLine = iterator.next();
        initHeaderAlias(firstLine);
        InnerExcelUtil.checkColumnNum(headerAliasMap.size());
        writeHeadRow(headerAliasMap.values());
        writeRow(firstLine);

        // 处理后续行的信息
        while(iterator.hasNext()) {
            Object line = iterator.next();
            writeRow(line);
        }

        return this;
    }

    @Override
    public IExcelWriter write(Collection<Map<String, Object>> mapList, Class<?> targetClass) {
        List<?> objectList = convertMap2List(mapList, targetClass);
        return write(objectList);
    }

    /**
     * map 转换为数据库查询结果列表
     *
     * @param results 查询结果
     * @param clazz   对象信息
     * @return 结果
     */
    private List<Object> convertMap2List(Iterable<Map<String, Object>> results, final Class<?> clazz) {
        try {
            List<Field> fieldList = ClassUtil.getAllFieldList(clazz);
            List<Object> resultList = new ArrayList<>();

            for (Map<String, Object> result : results) {
                Object pojo = clazz.newInstance();
                for (Field field : fieldList) {
                    String fieldName = getFieldName(field);
                    Object fieldValue = result.get(fieldName);
                    // 对于基本类型默认值的处理 避免报错
                    if(fieldValue != null) {
                        field.set(pojo, fieldValue);
                    }
                }
                resultList.add(pojo);
            }

            return resultList;
        } catch (InstantiationException | IllegalAccessException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    /**
     * 获取字段名称
     *
     * @param field 字段
     * @return 字段名称
     */
    private String getFieldName(final Field field) {
        final String fieldName = field.getName();
        if (field.isAnnotationPresent(ExcelField.class)) {
            ExcelField column = field.getAnnotation(ExcelField.class);
            return fieldName;
        }
        return fieldName;
    }

    /**
     * 根据表头的顺序构建对应信息。
     * @param object 对象信息
     * @return 构建后的列表信息
     */
    private Iterable<?> buildRowValues(final Object object) {
        Map beanMap = BeanUtil.beanToMap(object);
        List<Object> valueList = new ArrayList<>();
        for(String fieldName : headerAliasMap.keySet()) {
            Object value = beanMap.get(fieldName);
            valueList.add(value);
        }
        return valueList;
    }

    @Override
    public IExcelWriter flush(OutputStream outputStream) {
        try {
            this.workbook.write(outputStream);
            return this;
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    @Override
    public void close() {
        // 清空对象
        this.headerAliasMap = null;
        this.currentRow = null;
        this.isClosed = true;

        this.styleSet = null;
        this.workbook = null;
        this.sheet = null;
    }

    /**
     * 写 excel 表头
     * 获取 head 真正的信息
     * 如果没有，则视为没有字段需要写入。
     */
    private void writeHeadRow(final Iterable<?> headRowData) {
        if(!containsHeadRow) {
            InnerExcelUtil.writeRow(this.sheet.createRow(this.currentRow.getAndIncrement()),
                    headRowData, this.styleSet,
                    true);
            containsHeadRow = true;
        }
    }

    /**
     * 写出一行数据<br>
     * 本方法只是将数据写入Workbook中的Sheet，并不写出到文件<br>
     *
     * @param object 一行的数据
     */
    private void writeRow(final Object object) {
        Iterable<?> rowData = buildRowValues(object);
        InnerExcelUtil.writeRow(this.sheet.createRow(this.currentRow.getAndIncrement()),
                rowData, this.styleSet,
                false);
    }

    /**
     * 初始化标题别名
     * （1）根据 {@link ExcelField#order()} 进行排序
     * （2）如果没指定属性值，则默认 order=0
     *
     * @since 0.0.5
     */
    private void initHeaderAlias(final Object object) {
        // 这里可以添加类型校验
        ArgUtil.notNull(object, "object");
        final Class beanClass = object.getClass();

        // 按照顺序构建 map 信息
        headerAliasMap = Instances.singleton(HeaderAliasCache.class)
                .get(beanClass);
    }

    /**
     * 检查关闭的状态信息
     */
    private void checkClosedStatus() {
        if(this.isClosed) {
            throw new ExcelRuntimeException("ExcelWriter has been closed!");
        }
    }

    /**
     * 校验行数
     * @param dataSize 数据大小
     */
    private void checkRowNum(int dataSize) {
        int limit = getMaxRowNumLimit();
        if(dataSize > limit) {
            throw new ExcelRuntimeException("超出最大行数限制");
        }
    }

}
