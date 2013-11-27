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
    public static final String JEXCEL_CONTEXT_KEY = "util";

    public static final int MAX_COLUMN_TO_READ_COMMENT = 50;

    Workbook workbook;
    WritableWorkbook writableWorkbook;

    public JexcelTransformer() {
    }

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

    @Override
    public Context createInitialContext() {
        Context context = new Context();
        context.putVar(JEXCEL_CONTEXT_KEY, new JexcelUtil());
        return context;
    }

    private void readCellData(){
        int numberOfSheets = workbook.getNumberOfSheets();
        for(int i = 0; i < numberOfSheets; i++){
            Sheet sheet = workbook.getSheet(i);
            SheetData sheetData = JexcelSheetData.createSheetData(sheet);
            sheetMap.put(sheetData.getSheetName(), sheetData);
        }
    }

    public WritableWorkbook getWritableWorkbook() {
        return writableWorkbook;
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
                JexcelUtil.copySheetProperties(workbook.getSheet( srcCellRef.getSheetName() ), destSheet);
            }
            SheetData sheetData = sheetMap.get(srcCellRef.getSheetName());
            if(!isIgnoreColumnProps()){
                CellView columnView = destSheet.getColumnView(targetCellRef.getCol());
                columnView.setSize(sheetData.getColumnWidth(srcCellRef.getCol()));
                destSheet.setColumnView(targetCellRef.getCol(), columnView);
            }
            if(!isIgnoreRowProps()){
                try {
                    CellView rowView = destSheet.getRowView(targetCellRef.getRow());
                    rowView.setSize(sheetData.getRowData(srcCellRef.getRow()).getHeight());
                    destSheet.setRowView(targetCellRef.getRow(), rowView);
                } catch (RowsExceededException e) {
                    logger.warn("Failed to set row height for " + targetCellRef.getCellName(), e);
                }
            }
            try{
                ((JexcelCellData)cellData).writeToCell(destSheet, targetCellRef.getCol(), targetCellRef.getRow(), context);
                copyMergedRegions(cellData, targetCellRef);
            }catch(Exception e){
                logger.error("Failed to write a cell with " + cellData + " and " + context, e);
            }
        }
    }

    private void copyMergedRegions(CellData sourceCellData, CellRef destCell) throws WriteException {
        if(sourceCellData.getSheetName() == null ){ throw new IllegalArgumentException("Sheet name is null in copyMergedRegions");}
        JexcelSheetData sheetData = (JexcelSheetData)sheetMap.get( sourceCellData.getSheetName() );
        Range cellMergedRegion = null;
        for (Range mergedRegion : sheetData.getMergedCells()) {
            if(mergedRegion.getTopLeft().getRow() == sourceCellData.getRow() && mergedRegion.getTopLeft().getColumn() == sourceCellData.getCol()){
                cellMergedRegion = mergedRegion;
                break;
            }
        }
        if( cellMergedRegion != null){
            findAndRemoveExistingCellRegion(destCell);
            WritableSheet destSheet = writableWorkbook.getSheet(destCell.getSheetName());
            destSheet.mergeCells(destCell.getCol(), destCell.getRow(),
                    destCell.getCol() + cellMergedRegion.getBottomRight().getColumn() - cellMergedRegion.getTopLeft().getColumn(),
                    destCell.getRow() + cellMergedRegion.getBottomRight().getRow() - cellMergedRegion.getTopLeft().getRow());
        }
    }

    private void findAndRemoveExistingCellRegion(CellRef cellRef) {
        WritableSheet destSheet = writableWorkbook.getSheet(cellRef.getSheetName());
        Range[] mergedRegions = destSheet.getMergedCells();
        for(Range mergedRegion : mergedRegions){
            if(mergedRegion.getTopLeft().getRow() <= cellRef.getRow() && mergedRegion.getBottomRight().getRow() >= cellRef.getRow() &&
                    mergedRegion.getTopLeft().getColumn() <= cellRef.getCol() && mergedRegion.getBottomRight().getColumn() >= cellRef.getCol() ){
                destSheet.unmergeCells( mergedRegion );
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
        Cell cell = sheet.getCell(cellRef.getCol(), cellRef.getRow());
        WritableCell writableCell = new Formula(cellRef.getCol(), cellRef.getRow(), formulaString);
        if(cell != null && cell.getCellFormat() != null){
            writableCell.setCellFormat(cell.getCellFormat());
        }
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

    public void write() throws IOException {
        if(writableWorkbook != null){
            writableWorkbook.write();
            try {
                writableWorkbook.close();
            } catch (WriteException e) {
                throw new IllegalStateException("Cannot close a writable workbook", e);
            }
        }else{
            throw new IllegalStateException("An attempt to write an output stream with an uninitialized WritableWorkbook");
        }
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
