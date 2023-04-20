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

    @Override
    public FormulaData clone() {
        FormulaData formulaData = new FormulaData();
        formulaData.setFormulaValue(getFormulaValue());
        formulaData.setSharedFormula(isSharedFormula());
        return formulaData;
    }
}
