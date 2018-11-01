package com.github.houbb.iexcel.analyser;

import com.github.houbb.iexcel.analyser.sax.BaseSaxAnalyser;
import com.github.houbb.iexcel.analyser.sax.XlsSaxAnalyser;
import com.github.houbb.iexcel.analyser.sax.XlsxSaxAnalyser;
import com.github.houbb.iexcel.constant.enums.ExcelTypeEnum;
import com.github.houbb.iexcel.context.AnalysisContext;
import com.github.houbb.iexcel.context.AnalysisContextImpl;
import com.github.houbb.iexcel.exception.ExcelRuntimeException;
import com.github.houbb.iexcel.metadata.Sheet;

import java.io.InputStream;
import java.util.List;

/**
 * @author jipengfei
 */
public class ExcelAnalyserImpl implements ExcelAnalyser {

    private AnalysisContext analysisContext;

    private BaseSaxAnalyser saxAnalyser;

    private BaseSaxAnalyser getSaxAnalyser() {
        if (saxAnalyser == null) {
            try {
                if (ExcelTypeEnum.XLS.equals(analysisContext.getExcelType())) {
                    this.saxAnalyser = new XlsSaxAnalyser(analysisContext);
                } else {
                    this.saxAnalyser = new XlsxSaxAnalyser(analysisContext);
                }
            } catch (Exception e) {
                throw new ExcelRuntimeException("Analyse excel occur file error fileType " + analysisContext.getExcelType(),
                    e);
            }
        }
        return this.saxAnalyser;
    }

    @Override
    public void init(InputStream inputStream, ExcelTypeEnum excelTypeEnum, Object custom,
                     boolean trim) {
        analysisContext = new AnalysisContextImpl(inputStream, excelTypeEnum, custom,
            trim);
    }

    @Override
    public void analysis(Sheet sheetParam) {
        analysisContext.setCurrentSheet(sheetParam);
        analysis();
    }

    @Override
    public void analysis() {
        BaseSaxAnalyser saxAnalyser = getSaxAnalyser();
        saxAnalyser.execute();
    }

    @Override
    public List<Sheet> getSheets() {
        BaseSaxAnalyser saxAnalyser = getSaxAnalyser();
        return saxAnalyser.getSheets();
    }

}
