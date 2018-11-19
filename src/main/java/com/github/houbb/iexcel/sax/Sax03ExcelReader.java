package com.github.houbb.iexcel.sax;

import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.sax.handler.SaxRowHandler;
import com.github.houbb.iexcel.sax.handler.SaxRowHandlerContext;
import com.github.houbb.iexcel.sax.handler.impl.DefaultSaxRowHandler;
import com.github.houbb.iexcel.util.StrUtil;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 参考资料：[POI Sax 事件驱动解析Excel2003文件](http://www.cnblogs.com/wshsdlau/p/5643862.html)
 *
 * @author binbin.hou
 * @date 2018/11/16 13:53
 */
public class Sax03ExcelReader<T> extends AbstractSaxExcelReader<T> implements HSSFListener {

    //region 私有变量
    /**
     * 如果为公式，true表示输出公式计算后的结果值，false表示输出公式本身
     */
    private boolean isOutputFormulaValues = true;

    /**
     * 用于解析公式
     */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;

    /**
     * 子工作簿，用于公式计算
     */
    private HSSFWorkbook stubWorkbook;

    /**
     * 静态字符串表
     */
    private SSTRecord sstRecord;

    /**
     * 代理HSSFListener，用于跟踪文档格式记录，并提供一种简单的方法来查找单元格从其ID中使用的格式字符串。
     */
    private FormatTrackingHSSFListener formatListener;

    /**
     * Sheet边界记录，此Record中可以获得Sheet名
     */
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    private boolean isOutputNextStringRecord;

    /**
     * 存储一行所有列记录的容器
     */
    private List<Object> rowCellList = new ArrayList<>();

    /**
     * 用于存放最后的对象转换结果
     */
    private List<T> rowResultList = new ArrayList<>();

    /**
     * 当前表索引
     */
    private int curSheetIndex = -1;

    /**
     * 目标类型
     */
    private Class<T> targetClass;

    /**
     * 开始的行索引
     */
    private int startRowIndex;

    /**
     * 结束的行索引
     */
    private int endRowIndex;

    /**
     * 当前 excel 的列信息和待转换类 Field 之间的关系
     */
    private Map<Integer, Field> indexFieldMap = new HashMap<>();

    /**
     * 行处理器
     */
    private SaxRowHandler rowHandler = new DefaultSaxRowHandler();
    //endregion

    //region 对象构建
    public Sax03ExcelReader(File excelFile) {
        super(excelFile);
    }

    public Sax03ExcelReader(File excelFile, int sheetIndex) {
        super(excelFile, sheetIndex);
    }
    //endregion

    /**
     * HSSFListener 监听方法，处理 Record
     *
     * @param record 记录
     */
    @Override
    public void processRecord(Record record) {
        if (this.sheetIndex > -1 && this.curSheetIndex > this.sheetIndex) {
            // 指定Sheet之后的数据不再处理
            return;
        }

        if (record instanceof BoundSheetRecord) {
            // Sheet边界记录，此Record中可以获得Sheet名
            boundSheetRecords.add((BoundSheetRecord) record);
        } else if (record instanceof SSTRecord) {
            // 静态字符串表
            sstRecord = (SSTRecord) record;
        } else if (record instanceof BOFRecord) {
            BOFRecord bofRecord = (BOFRecord) record;
            if (bofRecord.getType() == BOFRecord.TYPE_WORKSHEET) {
                // 如果有需要，则建立子工作薄
                if (workbookBuildingListener != null && stubWorkbook == null) {
                    stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                }

                curSheetIndex++;
            }
        } else if (isProcessCurrentSheet()) {
            if (record instanceof MissingCellDummyRecord) {
                // 空值的操作
                MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
                rowCellList.add(mc.getColumn(), StrUtil.EMPTY);
            } else if (record instanceof LastCellOfRowDummyRecord) {
                // 行结束
                processLastCell((LastCellOfRowDummyRecord) record);
            } else {
                // 处理单元格值
                processCellValue(record);
            }
        }

    }

    /**
     * 处理单元格值
     *
     * @param record 单元格
     */
    private void processCellValue(Record record) {
        Object value = null;

        switch (record.getSid()) {
            case BlankRecord.sid:
                // 空白记录
                BlankRecord brec = (BlankRecord) record;
                rowCellList.add(brec.getColumn(), StrUtil.EMPTY);
                break;
            // 布尔类型
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                rowCellList.add(berec.getColumn(), berec.getBooleanValue());
                break;
            // 公式类型
            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;
                if (isOutputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        isOutputNextStringRecord = true;
                    } else {
                        value = formatListener.formatNumberDateCell(frec);
                    }
                } else {
                    value = StrUtil.DOUBLE_QUOTE +
                            HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression())
                            + StrUtil.DOUBLE_QUOTE;
                }
                rowCellList.add(frec.getColumn(), value);
                break;
            // 单元格中公式的字符串
            case StringRecord.sid:
                if (isOutputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    value = srec.getString();
                    isOutputNextStringRecord = false;
                }
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                this.rowCellList.add(lrec.getColumn(), value);
                break;
            // 字符串类型
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                if (sstRecord == null) {
                    rowCellList.add(lsrec.getColumn(), StrUtil.EMPTY);
                } else {
                    value = sstRecord.getString(lsrec.getSSTIndex()).toString();
                    rowCellList.add(lsrec.getColumn(), value);
                }
                break;
            // 数字类型
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                final String formatString = formatListener.getFormatString(numrec);
                if (formatString.contains(StrUtil.DOT)) {
                    //浮点数
                    value = numrec.getValue();
                } else if (formatString.contains(StrUtil.SLASH) || formatString.contains(StrUtil.COLON)) {
                    //日期
                    value = formatListener.formatNumberDateCell(numrec);
                } else {
                    double numValue = numrec.getValue();
                    final long longPart = (long) numValue;
                    // 对于无小数部分的数字类型，转为Long，否则保留原数字
                    value = (longPart == numValue) ? longPart : numValue;
                }

                // 向容器加入列值
                rowCellList.add(numrec.getColumn(), value);
                break;
            default:
                break;
        }
    }

    /**
     * 处理行结束后的操作，{@link LastCellOfRowDummyRecord}是行结束的标识Record
     * 1. 第一行处理表头相关信息
     * @param lastCell 行结束的标识Record
     */
    private void processLastCell(LastCellOfRowDummyRecord lastCell) {
        // 每行结束时， 调用handle() 方法
        int rowIndex = lastCell.getRow();

        // 当时开始的第一行且包含表头时，对于表头的处理
        // 如果是表头，则不再参与 list 的构建
        if(this.startRowIndex == rowIndex
            && containsHead) {
            this.indexFieldMap = InnerExcelUtil.getReadIndexFieldMap(this.targetClass, this.rowCellList);
            clear();
            return;
        }

        // 处理的范围判断
        if(startRowIndex <= rowIndex &&
            rowIndex <= endRowIndex) {
            SaxRowHandlerContext<T> context = new SaxRowHandlerContext<>();
            context.setTargetClass(this.targetClass);
            context.setRowIndex(rowIndex);
            context.setCellList(rowCellList);
            context.setIndexFieldMap(indexFieldMap);

            T row = this.rowHandler.handle(context);
            this.rowResultList.add(row);
        }

        clear();
    }

    private void clear() {
        // 清空行Cache
        this.rowCellList.clear();
    }

    /**
     * 是否处理当前sheet
     *
     * @return 是否处理当前sheet
     */
    private boolean isProcessCurrentSheet() {
        return this.sheetIndex < 0 || this.curSheetIndex == this.sheetIndex;
    }
    // ---------------------------------------------------------------------------------------------- Private method end

    @Override
    public List<T>  read(Class<T> tClass, int startIndex, int endIndex) {
        try {
            this.targetClass = tClass;
            this.startRowIndex = startIndex;
            this.endRowIndex = endIndex;

            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(excelFile));
            formatListener = new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(this));
            final HSSFRequest request = new HSSFRequest();
            if (isOutputFormulaValues) {
                request.addListenerForAllRecords(formatListener);
            } else {
                workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
                request.addListenerForAllRecords(workbookBuildingListener);
            }
            final HSSFEventFactory factory = new HSSFEventFactory();
            factory.processWorkbookEvents(request, fs);
        } catch (IOException e) {
            throw new ExcelRuntimeException(e);
        }

        return rowResultList;
    }

}
