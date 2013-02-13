package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.CellData;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import jxl.*;
import jxl.biff.formula.FormulaException;
import jxl.format.CellFormat;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author Leonid Vysochyn
 */
public class JexcelCellData extends CellData {
    static Logger logger = LoggerFactory.getLogger(JexcelCellData.class);

    jxl.CellType jxlCellType;
    CellFormat cellFormat;
    CellFeatures cellFeatures;

    public JexcelCellData(CellRef cellRef) {
        super(cellRef);
    }

    public static JexcelCellData createCellData(CellRef cellRef, Cell cell){
        JexcelCellData cellData = new JexcelCellData(cellRef);
        cellData.readCell(cell);
        cellData.updateFormulaValue();
        return cellData;
    }

    private void readCell(Cell cell) {
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        cellFeatures = cell.getCellFeatures();
        String comment = cellFeatures.getComment();
        if(comment != null){
            setCellComment(comment);
        }
    }

    private void readCellContents(Cell cell) {
        jxlCellType = cell.getType();
        try {
            if (jxlCellType == jxl.CellType.LABEL){
                LabelCell labelCell = (LabelCell) cell;
                cellValue = labelCell.getString();
                cellType = CellType.STRING;
            }else if (jxlCellType == jxl.CellType.NUMBER){
                NumberCell numberCell = (NumberCell) cell;
                cellValue = numberCell.getValue();
                cellType = CellType.NUMBER;
            }else if (jxlCellType == jxl.CellType.DATE){
                DateCell dateCell = (DateCell) cell;
                cellValue = dateCell.getDate();
                cellType = CellType.DATE;
            }else if (jxlCellType == jxl.CellType.BOOLEAN){
                BooleanCell booleanCell = (BooleanCell) cell;
                cellValue = booleanCell.getValue();
                cellType = CellType.BOOLEAN;
            }else if (jxlCellType == jxl.CellType.STRING_FORMULA || jxlCellType == jxl.CellType.BOOLEAN_FORMULA ||
                    jxlCellType == jxl.CellType.DATE_FORMULA || jxlCellType == jxl.CellType.NUMBER_FORMULA){
                FormulaCell formulaCell = (FormulaCell) cell;
                formula = formulaCell.getFormula();
                cellValue = formula;
                cellType = CellType.FORMULA;
            }else if (jxlCellType == jxl.CellType.ERROR || jxlCellType == jxl.CellType.FORMULA_ERROR){
                ErrorCell errorCell = (ErrorCell) cell;
                cellValue = errorCell.getErrorCode();
                cellType = CellType.ERROR;
            }else if ( jxlCellType == jxl.CellType.EMPTY ){
                cellValue = null;
                cellType = CellType.BLANK;
            }
            evaluationResult = cellValue;
        } catch (FormulaException e) {
            logger.warn("Failed to read formula", e);
        }
    }

    private void readCellStyle(Cell cell) {
        cellFormat = cell.getCellFormat();
    }

    public WritableCell createWritableCell(WritableSheet sheet, int col, int row, Context context){

    }

    public void writeToCell(WritableCell cell, Context context){
        evaluate(context);
        updateCellGeneralInfo(cell);
        updateCellContents( cell );
        updateCellStyle( cell );
    }

    private void updateCellGeneralInfo(WritableCell cell) {
        cell.setCellType( getPoiCellType(targetCellType) );
        if(comment != null ){
            cell.setCellComment(comment);
        }
    }

    private void updateCellContents(WritableCell cell) {

    }

    private void updateCellStyle(WritableCell cell) {

    }

}
