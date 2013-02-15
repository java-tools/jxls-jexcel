package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.*;
import com.jxls.plus.transform.AbstractTransformer;
import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class JexcelTransformer extends AbstractTransformer {
    static Logger logger = LoggerFactory.getLogger(JexcelTransformer.class);

    public static final int MAX_COLUMN_TO_READ_COMMENT = 50;

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
            if(!isIgnoreRowProps()){
                try {
                    destSheet.setRowView(targetCellRef.getRow(), sheetData.getRowData(srcCellRef.getRow()).getHeight());
                } catch (RowsExceededException e) {
                    logger.warn("Failed to set row height for " + targetCellRef.getCellName(), e);
                }
            }
            try{
                ((JexcelCellData)cellData).writeToCell(destSheet, targetCellRef.getCol(), targetCellRef.getRow(), context);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }

    }

    public void setFormula(CellRef cellRef, String formulaString) {
        if(cellRef == null || cellRef.getSheetName() == null ) return;
        WritableSheet sheet = writableWorkbook.getSheet(cellRef.getSheetName());
        if( sheet == null){
            int numberOfSheets = writableWorkbook.getNumberOfSheets();
            sheet = writableWorkbook.createSheet(cellRef.getSheetName(), numberOfSheets);
        }
        WritableCell writableCell = new Formula(cellRef.getCol(), cellRef.getRow(), formulaString);
        try{
            sheet.addCell(writableCell);
        }catch (Exception e){
            logger.error("Failed to set formula = " + formulaString + " into cell = " + cellRef.getCellName(), e);
        }
    }

    public void clearCell(CellRef cellRef) {
        if(cellRef == null || cellRef.getSheetName() == null ) return;
        WritableSheet sheet = writableWorkbook.getSheet(cellRef.getSheetName());
        if( sheet == null ) return;
        Blank blankCell = new Blank(cellRef.getCol(), cellRef.getRow());
        try {
            sheet.addCell(blankCell);
        } catch (WriteException e) {
            logger.error("Failed to clean up cell " + cellRef.getCellName(), e);
        }
    }

    public List<CellData> getCommentedCells() {
        List<CellData> commentedCells = new ArrayList<CellData>();
        for (SheetData sheetData : sheetMap.values()) {
            for (RowData rowData : sheetData) {
                if( rowData == null ) continue;
                for (CellData cellData : rowData) {
                    if(cellData != null && cellData.getCellComment() != null ){
                        commentedCells.add(cellData);
                    }
                }
                if( rowData.getNumberOfCells() == 0 ){
                    List<CellData> commentedCellData = readCommentsFromSheet(((JexcelSheetData)sheetData).getSheet(),  ((JexcelRowData)rowData).getRow());
                    commentedCells.addAll( commentedCellData );
                }
            }
        }
        return commentedCells;
    }

    public void addImage(AreaRef areaRef, int imageIdx) {
    }

    public void addImage(AreaRef areaRef, byte[] imageBytes, ImageType imageType) {
        if( imageType == null ){
            throw new IllegalArgumentException("Image type is undefined");
        }
        if( imageType != ImageType.PNG){
            throw new IllegalArgumentException("Only PNG images are currently supported");
        }
        WritableSheet sheet = writableWorkbook.getSheet(areaRef.getSheetName());
        sheet.addImage(new WritableImage(areaRef.getFirstCellRef().getCol(),areaRef.getFirstCellRef().getRow(),
        areaRef.getLastCellRef().getCol() - areaRef.getFirstCellRef().getCol(),
        areaRef.getLastCellRef().getRow() - areaRef.getFirstCellRef().getRow(),imageBytes));
    }

    private List<CellData> readCommentsFromSheet(Sheet sheet, Cell[] cells) {
        List<CellData> commentDataCells = new ArrayList<CellData>();
        for (Cell cell : cells) {
            CellFeatures cellFeatures = cell.getCellFeatures();
            if (cellFeatures.getComment() != null) {
                CellData cellData = new CellData(new CellRef(sheet.getName(), cell.getRow(), cell.getColumn()));
                cellData.setCellComment(cellFeatures.getComment());
                commentDataCells.add(cellData);
            }
        }
        return commentDataCells;
    }
}
