package com.github.houbb.iexcel.analyser.sax;

import com.github.houbb.iexcel.constant.ExcelXmlConst;
import com.github.houbb.iexcel.constant.enums.FieldTypeEnum;
import com.github.houbb.iexcel.context.AnalysisContext;
import com.github.houbb.iexcel.util.PositionUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.Arrays;

import static com.github.houbb.iexcel.constant.ExcelXmlConst.*;

/**
 * @author jipengfei
 */
public class XlsxRowHandler extends DefaultHandler {

    private FieldTypeEnum currentCellType;

    private int curRow;

    private int curCol;

    private String[] curRowContent = new String[20];

    private String currentCellValue;

    private SharedStringsTable sst;

    private AnalysisContext analysisContext;

    public XlsxRowHandler(SharedStringsTable sst,
                          AnalysisContext analysisContext) {
        this.analysisContext = analysisContext;
        this.sst = sst;

    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {

        setTotalRowCount(name, attributes);

        startCell(name, attributes);

        startCellValue(name);

    }

    private void startCellValue(String name) {
        if (name.equals(CELL_VALUE_TAG) || name.equals(CELL_VALUE_TAG_1)) {
            // initialize current cell value
            currentCellValue = "";
        }
    }

    private void startCell(String name, Attributes attributes) {
        if (CELL_TAG.equals(name)) {
            String currentCellIndex = attributes.getValue(ExcelXmlConst.POSITION);
            int nextRow = PositionUtil.getRow(currentCellIndex);
            if (nextRow > curRow) {
                curRow = nextRow;
                // endRow(ROW_TAG);
            }
            analysisContext.setCurrentRowNum(curRow);
            curCol = PositionUtil.getCol(currentCellIndex);

            String cellType = attributes.getValue("t");
            currentCellType = FieldTypeEnum.EMPTY;
            if (cellType != null && cellType.equals("s")) {
                currentCellType = FieldTypeEnum.STRING;
            }
            //if ("6".equals(attributes.getValue("s"))) {
            //    // date
            //    currentCellType = FieldType.DATE;
            //}

        }
    }

    private void endCellValue(String name) throws SAXException {
        // ensure size
        if (curCol >= curRowContent.length) {
            curRowContent = Arrays.copyOf(curRowContent, (int)(curCol * 1.5));
        }
        if (CELL_VALUE_TAG.equals(name)) {

            switch (currentCellType) {
                case STRING:
                    int idx = Integer.parseInt(currentCellValue);
                    currentCellValue = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
                    currentCellType = FieldTypeEnum.EMPTY;
                    break;
                //case DATE:
                //    Date dateVal = HSSFDateUtil.getJavaDate(Double.parseDouble(currentCellValue),
                //        analysisContext.use1904WindowDate());
                //    currentCellValue = TypeUtil.getDefaultDateString(dateVal);
                //    currentCellType = FieldType.EMPTY;
                //    break;
            }
            curRowContent[curCol] = currentCellValue;
        } else if (CELL_VALUE_TAG_1.equals(name)) {
            curRowContent[curCol] = currentCellValue;
        }
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        endRow(name);
        endCellValue(name);
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        currentCellValue += new String(ch, start, length);
    }


    private void setTotalRowCount(String name, Attributes attributes) {
        if (DIMENSION.equals(name)) {
            String d = attributes.getValue(DIMENSION_REF);
            String totalStr = d.substring(d.indexOf(":") + 1, d.length());
            String c = totalStr.toUpperCase().replaceAll("[A-Z]", "");
            analysisContext.setTotalCount(Integer.parseInt(c));
        }

    }

    private void endRow(String name) {
        if (name.equals(ROW_TAG)) {
            curRowContent = new String[20];
        }
    }

}

