package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.CellData;
import com.jxls.plus.common.CellRef;
import com.jxls.plus.common.Context;
import com.jxls.plus.util.Util;
import jxl.*;
import jxl.Cell;
import jxl.biff.formula.FormulaException;
import jxl.format.CellFormat;
import jxl.write.*;
import jxl.write.Boolean;
import jxl.write.Number;
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

    public CellFormat getCellFormat() {
        return cellFormat;
    }

    private void readCell(Cell cell) {
        readCellGeneralInfo(cell);
        readCellContents(cell);
        readCellStyle(cell);
    }

    private void readCellGeneralInfo(Cell cell) {
        cellFeatures = cell.getCellFeatures();
        if( cellFeatures != null ){
            String comment = cellFeatures.getComment();
            if(comment != null){
                setCellComment(comment);
            }
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
                    jxlCellType == jxl.CellType.DATE_FORMULA || jxlCellType == jxl.CellType.NUMBER_FORMULA ||
                    jxlCellType == jxl.CellType.FORMULA_ERROR){
                FormulaCell formulaCell = (FormulaCell) cell;
                formula = formulaCell.getFormula();
                cellValue = formula;
                cellType = CellType.FORMULA;
            }else if(jxlCellType == jxl.CellType.ERROR && cell instanceof Formula){
                Formula formulaCell = (Formula) cell;
                formula = formulaCell.getContents();
                cellValue = formula;
                cellType = CellType.FORMULA;
            }else if (jxlCellType == jxl.CellType.ERROR){
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

    public void writeToCell(WritableSheet sheet, int col, int row, Context context) throws WriteException {
        evaluate(context);
        if( evaluationResult != null && evaluationResult instanceof WritableCellValue){
            WritableCell cell = ((WritableCellValue)evaluationResult).writeToCell(sheet, col, row, context);
            updateCellStyle(cell);
        }else{
            WritableCell writableCell = createWritableCell(col, row);
            updateCellGeneralInfo(writableCell);
            updateCellStyle( writableCell );
            sheet.addCell(writableCell);
        }

    }

    private WritableCell createWritableCell(int col, int row) {
        WritableCell writableCell = null;
        switch(targetCellType){
            case STRING:
                if( !(evaluationResult instanceof byte[])){
                    writableCell = new Label(col, row, (String) evaluationResult);
                }
                break;
            case BOOLEAN:
                writableCell = new Boolean(col, row, (java.lang.Boolean)evaluationResult );
                break;
            case NUMBER:
                double value = ((java.lang.Number)evaluationResult).doubleValue();
                writableCell = new Number(col, row, value);
                break;
            case FORMULA:
                if( Util.formulaContainsJointedCellRef((String) evaluationResult) ){
                    writableCell = new Label(col, row, (String) evaluationResult);
                }else{
                    writableCell = new Formula(col, row, (String) evaluationResult);
                }
                break;
            case ERROR:
                writableCell = new Blank(col, row);
                break;
            default:
                writableCell = new Blank(col, row);
                break;
        }
        return writableCell;
    }

    private void updateCellGeneralInfo(WritableCell cell) {
        if( cellFeatures != null ){
            WritableCellFeatures writableCellFeatures = new WritableCellFeatures(cellFeatures);
            if( JexcelUtil.isJxComment(getCellComment()) ){
                writableCellFeatures.removeComment();
            }
            cell.setCellFeatures(writableCellFeatures);
        }
    }

    private void updateCellStyle(WritableCell cell) {
//        WritableCellFormat writableCellFormat = new WritableCellFormat(cellFormat);
//        cell.setCellFormat(writableCellFormat);
        cell.setCellFormat(cellFormat);
    }

}
