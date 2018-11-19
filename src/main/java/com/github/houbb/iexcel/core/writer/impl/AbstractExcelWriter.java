package com.github.houbb.iexcel.core.writer.impl;


import com.github.houbb.iexcel.annotation.ExcelConverter;
import com.github.houbb.iexcel.annotation.ExcelField;
import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.core.conventer.IExcelConverter;
import com.github.houbb.iexcel.core.conventer.IExcelConverterFactory;
import com.github.houbb.iexcel.core.conventer.impl.DefaultExcelConverterFactory;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.style.StyleSet;
import com.github.houbb.iexcel.util.BeanUtil;
import com.github.houbb.iexcel.util.ClassUtil;
import com.github.houbb.iexcel.util.CollUtil;
import com.github.houbb.iexcel.util.StrUtil;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.commons.collections4.MapUtils;
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

    /**
     * 类型转换器工厂
     */
    private IExcelConverterFactory converterFactory;

    public AbstractExcelWriter() {
        this(null, null);
    }

    public AbstractExcelWriter(final String sheetName) {
        this(sheetName, null);
    }

    public AbstractExcelWriter(final String sheetName,
                               final IExcelConverterFactory converterFactory) {
        this.workbook = getWorkbook();
        // 使用默认的 excel 转换工厂
        if(converterFactory == null) {
            this.converterFactory = new DefaultExcelConverterFactory();
        }
        final String realSheetName = StrUtil.isBlank(sheetName)
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
        if(CollUtil.isEmpty(data)) {
            return this;
        }

        // 状态校验
        checkClosedStatus();
        checkRowNum(data.size());

        // 结果转换
        Collection<?> convertData = convertCollectionData(data);

        // 生成 excel 文件
        // 处理第一条数据信息
        Iterator iterator = convertData.iterator();
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

    /**
     * 转换器转换数据
     * 如果列表为空或者转换器不存在，则不进行转换。
     * @param originalData 原始数据列表
     * @return 转换后的结果
     */
    private Collection<?> convertCollectionData(Collection<?> originalData) {
        if(CollUtil.isEmpty(originalData)) {
            return originalData;
        }

        final Object firstElem = originalData.iterator().next();
        Optional<IExcelConverter> excelConverterOptional = getExcelConverter(firstElem);
        if(!excelConverterOptional.isPresent()) {
            return originalData;
        }

        IExcelConverter excelConverter = excelConverterOptional.get();
        return excelConverter.write(originalData);
    }

    /**
     * 获取对应的 excel 转换器
     * @param object 对象
     * @return excel 转换器
     */
    private Optional<IExcelConverter> getExcelConverter(Object object) {
        if(null == object) {
            return Optional.empty();
        }

        Class clazz = object.getClass();
        if (clazz.isAnnotationPresent(ExcelConverter.class)) {
            ExcelConverter excelConverter = (ExcelConverter) clazz.getAnnotation(ExcelConverter.class);
            Class<? extends IExcelConverter> excelConverterClass = excelConverter.value();
            IExcelConverter iExcelConverter = converterFactory.getConverter(excelConverterClass);
            if(iExcelConverter == null) {
                throw new ExcelRuntimeException("无法找到对应的转换器");
            }
            return Optional.of(iExcelConverter);
        }
        return Optional.empty();
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
            final String mapKey = column.mapKey();
            if (StrUtil.isNotBlank(mapKey)) {
                return mapKey;
            }
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

        this.converterFactory = null;
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
        InnerExcelUtil.writeRow(this.sheet.createRow(this.currentRow.getAndIncrement()),
                headRowData, this.styleSet,
                true);
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
     */
    private void initHeaderAlias(final Object object) {
        if(!BeanUtil.isBean(object.getClass())) {
            throw new ExcelRuntimeException("列表必须为 java Bean 对象列表");
        }

        // 一定要指定为有序的 Map
        headerAliasMap = new LinkedHashMap<>();
        Field[] fields = object.getClass().getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelField.class)) {
                ExcelField excel = field.getAnnotation(ExcelField.class);
                final String fieldName = field.getName();
                final String headName = InnerExcelUtil.getFieldHeadName(excel, field);
                if (excel.writeRequire()) {
                    headerAliasMap.put(fieldName, headName);
                }
            }
        }

        if(MapUtils.isEmpty(headerAliasMap)) {
            throw new ExcelRuntimeException("excel 表头信息为空");
        }
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
