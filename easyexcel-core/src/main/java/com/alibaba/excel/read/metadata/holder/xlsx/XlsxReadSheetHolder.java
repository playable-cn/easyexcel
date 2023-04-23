package com.alibaba.excel.read.metadata.holder.xlsx;

import java.util.Deque;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;

import com.alibaba.excel.metadata.data.FormulaData;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.read.metadata.holder.ReadSheetHolder;
import com.alibaba.excel.read.metadata.holder.ReadWorkbookHolder;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.openxml4j.opc.PackageRelationshipCollection;

/**
 * sheet holder
 *
 * @author Jiaju Zhuang
 */
@Getter
@Setter
@EqualsAndHashCode
public class XlsxReadSheetHolder extends ReadSheetHolder {
    /**
     * Record the label of the current operation to prevent NPE.
     */
    private Deque<String> tagDeque;
    /**
     * Current Column
     */
    private Integer columnIndex;
    /**
     * Data for current label.
     */
    private StringBuilder tempData;
    /**
     * Formula for current label.
     */
    private StringBuilder tempFormula;
    /**
     * excel Relationship
     */
    private PackageRelationshipCollection packageRelationshipCollection;

    private Map<String, FormulaData> sharedFormulas;

    private boolean isSharedFormula;

    private boolean isOriginalSharedFormula;

    private String sharedFormulaIndex;

    public XlsxReadSheetHolder(ReadSheet readSheet, ReadWorkbookHolder readWorkbookHolder) {
        super(readSheet, readWorkbookHolder);
        this.tagDeque = new LinkedList<String>();
        packageRelationshipCollection
            = ((XlsxReadWorkbookHolder)readWorkbookHolder).getPackageRelationshipCollectionMap().get(
            readSheet.getSheetNo());
        sharedFormulas = new HashMap<>();
    }
}
