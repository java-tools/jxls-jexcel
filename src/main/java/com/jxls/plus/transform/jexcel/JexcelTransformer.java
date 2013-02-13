package main.java.com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.*;
import com.jxls.plus.transform.AbstractTransformer;
import com.jxls.plus.transform.jexcel.JexcelCellData;
import com.jxls.plus.transform.jexcel.JexcelSheetData;
import jxl.CellView;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.biff.RowsExceededException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class JexcelTransformer extends AbstractTransformer {
    static Logger logger = LoggerFactory.getLogger(JexcelTransformer.class);

    Workbook workbook;
    WritableWorkbook writableWorkbook;

    private JexcelTransformer(Workbook workbook, WritableWorkbook writableWorkbook) {
        this.workbook = workbook;
        this.writableWorkbook = writableWorkbook;
    }

    public static JexcelTransformer createTransformer(InputStream is, OutputStream os) throws IOException, BiffException {
        Workbook workbook = Workbook.getWorkbook(is);
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(os, workbook);
        JexcelTransformer transformer = new JexcelTransformer(workbook, writableWorkbook);
        transformer.readCellData();
        return transformer;
    }

    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        for(int i = 0; i < numberOfSheets; i++){
            Sheet sheet = workbook.getSheet(i);
            SheetData sheetData = JexcelSheetData.createSheetData(sheet);
            sheetMap.put(sheetData.getSheetName(), sheetData);
        }
    }

    public void transform(CellRef srcCellRef, CellRef targetCellRef, Context context) {
        CellData cellData = this.getCellData(srcCellRef);
        if(cellData != null){
            cellData.addTargetPos(targetCellRef);
            if(targetCellRef == null || targetCellRef.getSheetName() == null){
                logger.info("Target cellRef is null or has empty sheet name, cellRef=" + targetCellRef);
                return;
            }
            WritableSheet destSheet = writableWorkbook.getSheet(targetCellRef.getSheetName());
            if(destSheet == null){
                int numberOfSheets = writableWorkbook.getNumberOfSheets();
                destSheet = writableWorkbook.createSheet(targetCellRef.getSheetName(), numberOfSheets);
            }
            SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
            if(!isIgnoreColumnProps()){
                destSheet.setColumnView(targetCellRef.getCol(), sheetData.getColumnWidth(srcCellRef.getCol()));
            }
            WritableCell destCell = destSheet.getWritableCell(targetCellRef.getCol(), targetCellRef.getRow());
            if(!isIgnoreRowProps()){
                try {
                    destSheet.setRowView(targetCellRef.getRow(), sheetData.getRowData(srcCellRef.getRow()).getHeight());
                } catch (RowsExceededException e) {
                    logger.warn("Failed to set row height for " + targetCellRef.getCellName());
                }
            }
            try{
                ((JexcelCellData)cellData).createWritableCell(destSheet, targetCellRef.getCol(), targetCellRef.getRow())
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }

    }

    public void setFormula(CellRef cellRef, String formulaString) {

    }

    public void clearCell(CellRef cellRef) {

    }

    public List<CellData> getCommentedCells() {
        return null;
    }

    public void addImage(AreaRef areaRef, int imageIdx) {
    }

    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType) {
    }
}
