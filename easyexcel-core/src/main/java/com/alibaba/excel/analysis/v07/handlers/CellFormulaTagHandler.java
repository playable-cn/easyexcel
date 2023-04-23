package com.alibaba.excel.analysis.v07.handlers;

import com.alibaba.excel.constant.ExcelXmlConstants;
import com.alibaba.excel.context.xlsx.XlsxReadContext;
import com.alibaba.excel.metadata.data.FormulaData;
import com.alibaba.excel.read.metadata.holder.xlsx.XlsxReadSheetHolder;

import org.xml.sax.Attributes;

import java.util.HashMap;
import java.util.Map;

/**
 * Cell Handler
 *
 * @author jipengfei
 */
public class CellFormulaTagHandler extends AbstractXlsxTagHandler {

    @Override
    public void startElement(XlsxReadContext xlsxReadContext, String name, Attributes attributes) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        xlsxReadSheetHolder.setTempFormula(new StringBuilder());
        xlsxReadSheetHolder.setSharedFormula(false);
        xlsxReadSheetHolder.setOriginalSharedFormula(false);
        xlsxReadSheetHolder.setSharedFormulaIndex("");

        String formulaType = attributes.getValue(ExcelXmlConstants.SHAREDSTRINGS_T_TAG);
        if (ExcelXmlConstants.SHARED_FORMULA_TYPE.equalsIgnoreCase(formulaType)) {
            String sharedFormulaIndex = attributes.getValue(ExcelXmlConstants.SHAREDSTRINGS_SI_TAG);
            xlsxReadSheetHolder.setSharedFormulaIndex(sharedFormulaIndex);
            Map<String, FormulaData> sharedFormulas = xlsxReadSheetHolder.getSharedFormulas();
            if (sharedFormulas.containsKey(sharedFormulaIndex)) {
                xlsxReadSheetHolder.setSharedFormula(true);
                FormulaData originalSharedFormulaData = sharedFormulas.get(sharedFormulaIndex);
                xlsxReadSheetHolder.getTempFormula().append(originalSharedFormulaData.getFormulaValue());
            } else {
                xlsxReadSheetHolder.setOriginalSharedFormula(true);
            }
        }
    }

    @Override
    public void endElement(XlsxReadContext xlsxReadContext, String name) {
        XlsxReadSheetHolder xlsxReadSheetHolder = xlsxReadContext.xlsxReadSheetHolder();
        FormulaData formulaData = new FormulaData();
        formulaData.setFormulaValue(xlsxReadSheetHolder.getTempFormula().toString());
        formulaData.setSharedFormula(xlsxReadSheetHolder.isSharedFormula());
        formulaData.setSharedRowIndex(xlsxReadSheetHolder.getRowIndex());
        formulaData.setSharedColumnIndex(xlsxReadSheetHolder.getColumnIndex());

        Map<String, FormulaData> sharedFormulas = xlsxReadSheetHolder.getSharedFormulas();
        if (xlsxReadSheetHolder.isOriginalSharedFormula()) {
            sharedFormulas.put(xlsxReadSheetHolder.getSharedFormulaIndex(), formulaData.clone());
        }

        if (xlsxReadSheetHolder.isSharedFormula()) {
            String sharedFormulaIndex = xlsxReadSheetHolder.getSharedFormulaIndex();
            if (sharedFormulas.containsKey(sharedFormulaIndex)) {
                FormulaData originalSharedFormulaData = sharedFormulas.get(sharedFormulaIndex);
                formulaData.setSharedRowIndex(originalSharedFormulaData.getSharedRowIndex());
                formulaData.setSharedColumnIndex(originalSharedFormulaData.getSharedColumnIndex());
            }
        }

        xlsxReadSheetHolder.getTempCellData().setFormulaData(formulaData);
    }

    @Override
    public void characters(XlsxReadContext xlsxReadContext, char[] ch, int start, int length) {
        xlsxReadContext.xlsxReadSheetHolder().getTempFormula().append(ch, start, length);
    }
}