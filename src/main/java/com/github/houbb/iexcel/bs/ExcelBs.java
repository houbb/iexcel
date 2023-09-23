package com.github.houbb.iexcel.bs;

import com.github.houbb.heaven.annotation.ThreadSafe;
import com.github.houbb.heaven.util.common.ArgUtil;
import com.github.houbb.heaven.util.guava.Guavas;
import com.github.houbb.heaven.util.io.FileUtil;
import com.github.houbb.heaven.util.util.CollectionUtil;
import com.github.houbb.iexcel.constant.ExcelConst;
import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.core.reader.IExcelReader;
import com.github.houbb.iexcel.core.writer.IExcelWriter;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.hutool.api.SaxReadHandler;
import com.github.houbb.iexcel.sax.AbstractSaxExcelReader;
import com.github.houbb.iexcel.sax.handler.AbstractSaxReadHandler;
import com.github.houbb.iexcel.util.excel.ExcelUtil;
import org.apache.poi.ss.formula.functions.T;

import java.io.*;
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
     * @since 0.0.4
     */
    private String path;

    /**
     * 大 excel 模式
     * @since 0.0.4
     */
    private boolean bigExcelMode = false;

    /**
     * 待写入列表
     * @since 0.0.4
     */
    private final List writeBufferList = Guavas.newArrayList();

    /**
     * 指定文件路径
     * @return 结果
     * @since 0.0.6
     * @see #path(String) 指定文件路径
     */
    public static ExcelBs newInstance() {
        return new ExcelBs();
    }

    /**
     * 指定文件路径
     * @param path 文件路径
     * @return 结果
     * @since 0.0.4
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
        ArgUtil.notEmpty(path, "path");
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
     * @since 0.0.4
     */
    @SuppressWarnings("unchecked")
    public ExcelBs append(final Collection<?> collection) {
        if(CollectionUtil.isEmpty(collection)) {
            return this;
        }

        this.writeBufferList.addAll(collection);
        return this;
    }

    /**
     * 写入当前缓存中的所有数据
     * @since 0.0.4
     */
    public void write() {
        ArgUtil.notEmpty(path, "path");

        // 获取对应的 IWrite
        IExcelWriter excelWriter = getExcelWriter(path);
        // 执行写入
        ExcelUtil.onceWriteAndFlush(excelWriter, writeBufferList, path);
    }

    /**
     * 获取文件输出流信息
     * 用户需要的流应该是可以直接 web 下载的文件流。
     * [java使用OutputStream实现下载文件示例](https://www.jianshu.com/p/065c4a5ae3a8)
     *
     * 也可以采用如下的思路：
     * （1）创建临时文件
     * （2）创建成功后获取对应文件字节信息
     * （3）删除临时文件。
     *
     * 因为直接返回 Stream 可能会导致流忘记关闭等问题。
     * @return 输出流
     * @since 0.0.6
     */
    public synchronized byte[] bytes() {
        // 获取对应的 IWrite
        final String tempFile = getTempFilePath();
        File file = new File(tempFile);
        try(IExcelWriter excelWriter = getExcelWriter(tempFile);
            OutputStream outputStream = new FileOutputStream(file)) {

            // 写入文件
            excelWriter.write(writeBufferList);
            excelWriter.flush(outputStream);

            return FileUtil.getFileBytes(file);
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        } finally {
            // 删除创建的文件
            FileUtil.deleteFile(file);
        }
    }

    /**
     * 获取临时文件路径
     * @return 文件路径
     * @since 0.0.6
     */
    private String getTempFilePath() {
        if(bigExcelMode) {
            return ExcelConst.INNER_EXCEL_TEMP_PREFIX+ExcelTypeEnum.XLSX.getValue();
        }
        return ExcelConst.INNER_EXCEL_TEMP_PREFIX+ExcelTypeEnum.XLS.getValue();
    }

    /**
     * 先写入缓存中的所有数据，然后写入当前集合中的所有数据。
     * @param collection 集合
     * @since 0.0.4
     */
    public void write(final Collection<?> collection) {
        this.append(collection).write();
    }

    /**
     * 获取对应的 excel 写入类
     * @return excel 写入类
     * @since 0.0.4
     */
    private IExcelWriter getExcelWriter(final String filePath) {
        IExcelWriter excelWriter;
        if(bigExcelMode) {
            excelWriter = ExcelUtil.getBigExcelWriter();
        } else {
            if(filePath.endsWith(ExcelTypeEnum.XLS.getValue())) {
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
     * @since 0.0.4
     */
    private IExcelReader getExcelReader() {
        ArgUtil.notEmpty(path, "path");

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
    @SuppressWarnings("unchecked")
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
    @SuppressWarnings("unchecked")
    public <T> List<T> read(Class<T> tClass, final int startIndex, final int endIndex) {
        IExcelReader excelReader = getExcelReader();
        return excelReader.read(tClass, startIndex, endIndex);
    }

    /**
     * sax 读
     * @param tClass 类
     * @param saxReadHandler 行处理类
     */
    public <T> void readBySax(final Class<T> tClass,
                              AbstractSaxReadHandler<T> saxReadHandler,
                              final int startIndex, final int endIndex) {
        bigExcelMode = true;

        AbstractSaxExcelReader<T> excelReader = (AbstractSaxExcelReader<T>) getExcelReader();
        excelReader.saxReadHandler(saxReadHandler);
        excelReader.read(tClass, startIndex, endIndex);
    }

}
