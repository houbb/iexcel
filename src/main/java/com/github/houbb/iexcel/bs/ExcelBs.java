package com.github.houbb.iexcel.bs;

import com.github.houbb.heaven.annotation.ThreadSafe;
import com.github.houbb.heaven.util.common.ArgUtil;
import com.github.houbb.heaven.util.guava.Guavas;
import com.github.houbb.heaven.util.util.CollectionUtil;
import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.util.excel.ExcelUtil;

import java.io.File;
import java.util.Collection;
import java.util.List;

/**
 * Excel Bs
 * 优化调用体验，屏蔽实现细节。
 * @author binbin.hou
 * @since 0.0.4
 */
@ThreadSafe
public final class ExcelBs {

    private ExcelBs(){}

    /**
     * 文件路径
     */
    private String path;

    /**
     * 大 excel 模式
     */
    private boolean bigExcelMode = false;

    /**
     * 待写入列表
     */
    private List writeBufferList = Guavas.newArrayList();

    /**
     * 指定编码
     * @param path 文件路径
     * @return 结果
     */
    public static ExcelBs newInstance(final String path) {
        ExcelBs excelBs = new ExcelBs();
        return excelBs.path(path);
    }

    /**
     * 设置文件路径
     * @param path 文件路径
     * @return this
     * @since 0.0.4
     */
    public ExcelBs path(final String path) {
        ArgUtil.notNull(path, "path");
        this.path = path;
        return this;
    }

    public String path() {
        return path;
    }

    public boolean bigWriteMode() {
        return bigExcelMode;
    }

    public ExcelBs bigWriteMode(boolean bigWriteMode) {
        this.bigExcelMode = bigWriteMode;
        return this;
    }

    //------------------------------------- 写入
    /**
     * 在写入列表中新增一个对象对象集合
     * @param collection 对象集合
     * @return this
     */
    public ExcelBs append(final Collection<?> collection) {
        if(CollectionUtil.isEmpty(collection)) {
            return this;
        }

        for(Object object : collection) {
            this.writeBufferList.add(object);
        }
        return this;
    }

    /**
     * 写入当前缓存中的所有数据
     */
    public void write() {
        // 获取对应的 IWrite
        IExcelWriter excelWriter = getExcelWriter();
        // 执行写入
        ExcelUtil.onceWriteAndFlush(excelWriter, writeBufferList, path);
    }

    /**
     * 先写入缓存中的所有数据，然后写入当前集合中的所有数据。
     * @param collection 集合
     */
    public void write(final Collection<?> collection) {
        this.append(collection).write();
    }

    /**
     * 获取对应的 excel 写入类
     * @return excel 写入类
     * @since 0.0.4
     */
    private IExcelWriter getExcelWriter() {
        IExcelWriter excelWriter;
        if(bigExcelMode) {
            excelWriter = ExcelUtil.getBigExcelWriter();
        } else {
            if(path.endsWith(ExcelTypeEnum.XLS.getValue())) {
                excelWriter = ExcelUtil.get03ExcelWriter();
            } else {
                excelWriter = ExcelUtil.get07ExcelWriter();
            }
        }
        return excelWriter;
    }

    //------------------------------------------------- 读取

    /**
     * 获取 excel 读取类
     * TODO: 后续考虑 index/name 之间的关系，保证二者不冲突。
     * 考虑可以使用 index/name 指定读取的 sheet。
     * @return 读取类实现
     */
    private IExcelReader getExcelReader() {
        File file = new File(path);
        if(!bigExcelMode) {
            return ExcelUtil.getExcelReader(file);
        } else {
            return ExcelUtil.getBigExcelReader(file);
        }
    }

    /**
     * 读取当前 sheet 的所有信息
     * @param tClass 对应的 javabean 类型
     * @return 对象列表
     * @since 0.0.4
     * @param <T> 泛型
     */
    public <T> List<T> read(Class<T> tClass) {
        IExcelReader excelReader = getExcelReader();
        return excelReader.readAll(tClass);
    }

    /**
     * 读取指定范围内的
     * @param tClass 泛型
     * @param startIndex 开始的行信息(从0开始) 是不包含 head 行的。
     * @param endIndex 结束的行信息
     * @return 读取的对象列表
     * @since 0.0.4
     * @param <T> 泛型
     */
    public <T> List<T> read(Class<T> tClass, final int startIndex, final int endIndex) {
        IExcelReader excelReader = getExcelReader();
        return excelReader.read(tClass, startIndex, endIndex);
    }

}
