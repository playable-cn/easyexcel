package com.alibaba.excel.metadata.data;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;

/**
 * formula
 *
 * @author Jiaju Zhuang
 */
@Getter
@Setter
@EqualsAndHashCode
public class FormulaData {
    /**
     * formula
     */
    private String formulaValue;

    /**
     * is shared formula
     */
    private boolean isSharedFormula;

    /**
     * shared formula cell row index
     */
    private Integer sharedRowIndex;

    /**
     * shared formula cell column index
     */
    private Integer sharedColumnIndex;

    @Override
    public FormulaData clone() {
        FormulaData formulaData = new FormulaData();
        formulaData.setFormulaValue(getFormulaValue());
        formulaData.setSharedFormula(isSharedFormula());
        formulaData.setSharedRowIndex(getSharedRowIndex());
        formulaData.setSharedColumnIndex(getSharedColumnIndex());
        return formulaData;
    }
}
