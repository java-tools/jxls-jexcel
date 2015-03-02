package org.jxls.transform.jexcel;

import org.jxls.common.CellData;
import org.jxls.common.CellRef;
import org.jxls.common.RowData;
import jxl.Cell;
import jxl.CellView;
import jxl.Sheet;

/**
 * @author Leonid Vysochyn
 */
public class JexcelRowData extends RowData {
    Cell[] cells;

    public static RowData createRowData(Sheet sheet, int rowNum){
        return createRowData(sheet.getName(), sheet.getRow(rowNum), sheet.getRowView(rowNum));
    }

    public static RowData createRowData(String sheetName, Cell[] cells, CellView rowCellView){
        if( cells == null ) return null;
        JexcelRowData rowData = new JexcelRowData();
        rowData.cells = cells;
        rowData.height = rowCellView.getSize();
        int numberOfCells = cells.length;
        for(int cellIndex = 0; cellIndex < numberOfCells; cellIndex++){
            Cell cell = cells[cellIndex];
            if(cell != null ){
                CellData cellData = JexcelCellData.createCellData(new CellRef(sheetName, cell.getRow(), cellIndex), cell);
                rowData.addCellData(cellData);
            }else{
                rowData.addCellData(null);
            }
        }
        return rowData;
    }

    public Cell[] getRow() {
        return cells;
    }
}
