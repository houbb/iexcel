package com.github.houbb.iexcel.sax;

import com.github.houbb.heaven.constant.PunctuationConst;
import com.github.houbb.heaven.util.lang.StringUtil;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.sax.handler.SaxRowHandler;
import com.github.houbb.iexcel.sax.handler.SaxRowHandlerContext;
import com.github.houbb.iexcel.sax.handler.impl.DefaultSaxRowHandler;
import com.github.houbb.iexcel.sax.util.CellDataType;
import com.github.houbb.iexcel.sax.util.ExcelSaxUtil;
import com.github.houbb.iexcel.util.excel.InnerExcelUtil;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.*;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.github.houbb.iexcel.sax.constant.Sax07Constant.*;

/**
 * excel 的 2007 的 sax 解析模式
 * 参考资料：
 * Excel2007格式说明见：http://www.cnblogs.com/wangmingshun/p/6654143.html
 * Default 处理器：http://www.cnblogs.com/yaozhenfa/p/np_android_DefaultHandler.html
 *
 * http://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 * @author binbin.hou
 * date 2018/11/16 13:53
 */
public class Sax07ExcelReader<T> extends AbstractSaxExcelReader<T> implements ContentHandler {

    //region 私有变量
    /**
     * 存储每行的列元素
     */
    private List<Object> rowCellList = new ArrayList<>();

    /**
     * 行数据列表
     */
    private List<T> rowResultList = new ArrayList<>();

    /**
     * 行处理器
     */
    private SaxRowHandler rowHandler = new DefaultSaxRowHandler();

    /**
     * 目标类型
     */
    private Class<T> targetClass;

    /**
     * 当前 excel 的列信息和待转换类 Field 之间的关系
     */
    private Map<Integer, Field> indexFieldMap = new HashMap<>();

    /**
     * 开始行
     */
    private int startRowIndex;

    /**
     * 结束行
     */
    private int endRowIndex;

    /**
     * 当前行
     */
    private int currentRow;

    /**
     * 当前列
     */
    private int currentCell;

    /**
     * 上一次的内容
     */
    private String lastContent;

    /**
     * 单元数据类型
     */
    private CellDataType cellDataType;

    /**
     * 当前列坐标， 如A1，B5
     */
    private String curCoordinate;

    /**
     * 前一个列的坐标
     */
    private String preCoordinate;

    /**
     * 行的最大列坐标
     */
    private String maxCellCoordinate;

    /**
     * excel 2007 的共享字符串表,对应 sharedString.xml
     */
    private SharedStringsTable sharedStringsTable;

    /**
     * 单元格的格式表，对应style.xml
     */
    private StylesTable stylesTable;

    /**
     * 单元格存储的格式化字符串，numFormat 的 formatCode属性的值
     */
    private String numFormatString;


    //endregion

    //region 对象构建开始
    public Sax07ExcelReader(File excelFile) {
        super(excelFile);
    }

    public Sax07ExcelReader(File excelFile, int sheetIndex) {
        super(excelFile, sheetIndex);
    }
    //endregion


    @Override
    public List<T> read(Class<T> tClass, int startIndex, int endIndex) {
        try(OPCPackage opcPackage = OPCPackage.open(this.excelFile)) {
            this.startRowIndex = startIndex;
            this.endRowIndex = endIndex;
            this.targetClass = tClass;


            final XSSFReader xssfReader = new XSSFReader(opcPackage);
            // 根据 rId# 或 rSheet# 查找sheet
            try(InputStream sheetInputStream = xssfReader.getSheet(RID_PREFIX + (sheetIndex + 1));){
                // 获取共享样式表
                stylesTable = xssfReader.getStylesTable();
                // 获取共享字符串表
                this.sharedStringsTable = xssfReader.getSharedStringsTable();
                parse(sheetInputStream);
                return this.rowResultList;
            } catch (SAXException e) {
                throw new ExcelRuntimeException(e);
            }
        } catch (IOException | OpenXML4JException e) {
            throw new ExcelRuntimeException(e);
        }
    }

    /**
     * 读到一个xml开始标签时的回调处理方法
     */
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
        // 单元格元素
        if (C_ELEMENT.equals(qName)) {
            // 获取当前列坐标
            String tempCurCoordinate = attributes.getValue(R_ATTR);
            // 前一列为null，则将其设置为"@",A为第一列，ascii码为65，前一列即为@，ascii码64
            if (preCoordinate == null) {
                preCoordinate = PunctuationConst.AT;
            } else {
                // 存在，则前一列要设置为上一列的坐标
                preCoordinate = curCoordinate;
            }
            // 重置当前列
            curCoordinate = tempCurCoordinate;
            // 设置单元格类型
            setCellType(attributes);
        }

        lastContent = PunctuationConst.EMPTY;
    }

    /**
     * 设置单元格的类型
     *
     * @param attribute 属性
     */
    private void setCellType(Attributes attribute) {
        // 重置numFmtIndex,numFmtString的值
        // 单元格存储格式的索引，对应style.xml中的numFmts元素的子元素索引
        int numFmtIndex = 0;
        numFormatString = PunctuationConst.EMPTY;
        this.cellDataType = CellDataType.of(attribute.getValue(T_ATTR_VALUE));

        // 获取单元格的xf索引，对应style.xml中cellXfs的子元素xf
        final String xfIndexStr = attribute.getValue(S_ATTR_VALUE);
        if (xfIndexStr != null) {
            int xfIndex = Integer.parseInt(xfIndexStr);
            XSSFCellStyle xssfCellStyle = stylesTable.getStyleAt(xfIndex);
            numFmtIndex = xssfCellStyle.getDataFormat();
            numFormatString = xssfCellStyle.getDataFormatString();

            if (numFormatString == null) {
                cellDataType = CellDataType.NULL;
                numFormatString = BuiltinFormats.getBuiltinFormat(numFmtIndex);
            } else if (org.apache.poi.ss.usermodel.DateUtil.isADateFormat(numFmtIndex, numFormatString)) {
                cellDataType = CellDataType.DATE;
            }
        }
    }

    /**
     * 标签结束的回调处理方法
     */
    @Override
    public void endElement(String uri, String localName, String qName) throws SAXException {
        final String contentStr = StringUtil.trim(lastContent);

        if (T_ELEMENT.equals(qName)) {
            // type标签
            rowCellList.add(currentCell++, contentStr);
        } else if (C_ELEMENT.equals(qName)) {
            // cell标签
            Object value = ExcelSaxUtil.getDataValue(this.cellDataType, contentStr,
                    this.sharedStringsTable, this.numFormatString);
            // 补全单元格之间的空格
            fillBlankCell(preCoordinate, curCoordinate, false);
            rowCellList.add(currentCell++, value);
        } else if (ROW_ELEMENT.equals(qName)) {
            // 如果是row标签，说明已经到了一行的结尾
            // 最大列坐标以第一行(开始行)的为准
            // 第一行处理 处理表头相关信息
            if (currentRow == startRowIndex) {
                maxCellCoordinate = curCoordinate;
            }

            // 补全一行尾部可能缺失的单元格
            if (maxCellCoordinate != null) {
                fillBlankCell(curCoordinate, maxCellCoordinate, true);
            }

            // 表头信息当前行判断
            if (currentRow == startRowIndex
                    && containsHead) {
                this.indexFieldMap = InnerExcelUtil.getReadIndexFieldMap(this.targetClass, this.rowCellList);
                clear();
                return;
            }

            if (currentRow >= startRowIndex
                    && currentRow <= endRowIndex) {
                SaxRowHandlerContext<T> context = new SaxRowHandlerContext<>();
                context.setCellList(this.rowCellList);
                context.setRowIndex(currentRow);
                context.setIndexFieldMap(this.indexFieldMap);
                context.setTargetClass(this.targetClass);
                T row = rowHandler.handle(context);
                this.rowResultList.add(row);
                clear();
            }
        }
    }

    /**
     * 结果置空
     */
    private void clear() {
        // 一行结束
        // 清空rowCellList,
        rowCellList.clear();
        // 行数增加
        currentRow++;
        // 当前列置0
        currentCell = 0;
        // 置空当前列坐标和前一列坐标
        curCoordinate = null;
        preCoordinate = null;
    }

    /**
     * s标签结束的回调处理方法
     */
    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        // 得到单元格内容的值
        lastContent = lastContent.concat(new String(ch, start, length));
    }

    @Override
    public void setDocumentLocator(Locator locator) {
        // pass
    }

    /**
     * ?xml标签的回调处理方法
     */
    @Override
    public void startDocument() throws SAXException {
        // pass
    }

    @Override
    public void endDocument() throws SAXException {
        // pass
    }

    @Override
    public void startPrefixMapping(String prefix, String uri) throws SAXException {
        // pass
    }

    @Override
    public void endPrefixMapping(String prefix) throws SAXException {
        // pass
    }

    @Override
    public void ignorableWhitespace(char[] ch, int start, int length) throws SAXException {
        // pass
    }

    @Override
    public void processingInstruction(String target, String data) throws SAXException {
        // pass
    }

    @Override
    public void skippedEntity(String name) throws SAXException {
        // pass
    }

    /**
     * 处理流中的Excel数据
     *
     * @param sheetInputStream sheet流
     * @throws IOException  IO异常
     * @throws SAXException SAX异常
     */
    private void parse(InputStream sheetInputStream) throws IOException, SAXException {
        fetchSheetReader().parse(new InputSource(sheetInputStream));
    }

    /**
     * 填充空白单元格，如果前一个单元格大于后一个，不需要填充<br>
     *
     * @param preCoordinate 前一个单元格坐标
     * @param curCoordinate 当前单元格坐标
     * @param isEnd         是否为最后一个单元格
     */
    private void fillBlankCell(String preCoordinate, String curCoordinate, boolean isEnd) {
        if (!curCoordinate.equals(preCoordinate)) {
            int len = ExcelSaxUtil.countNullCell(preCoordinate, curCoordinate);
            if (isEnd) {
                len++;
            }
            while (len-- > 0) {
                rowCellList.add(currentCell++, "");
            }
        }
    }

    /**
     * 获取sheet的解析器
     *
     * @return {@link XMLReader}
     * @throws SAXException SAX异常
     */
    private XMLReader fetchSheetReader() throws SAXException {
        XMLReader xmlReader = null;
        try {
            xmlReader = XMLReaderFactory.createXMLReader(CLASS_SAXPARSER);
        } catch (SAXException e) {
            if (e.getMessage().contains(CLASS_SAXPARSER)) {
                throw new ExcelRuntimeException("You need to add 'xerces:xercesImpl' to your project and version >= 2.11.0", e);
            } else {
                throw e;
            }
        }
        xmlReader.setContentHandler(this);
        return xmlReader;
    }
}
