package com.jin.commons.poi.model;

import org.apache.poi.ss.formula.Formula;

import java.util.Set;

/**
 * 公式
 */
public class FormulaSettings {

    private FormulaType formulaType;

    private Set<String> groupName;

    private FormulaSettings(){

    }

    public FormulaSettings(FormulaType formulaType){
        this.formulaType = formulaType;
    }

    public Set<String> getGroupName() {
        return groupName;
    }

    public void setGroupName(Set<String> groupName) {
        this.groupName = groupName;
    }

    public FormulaType getFormulaType() {
        return formulaType;
    }
}
