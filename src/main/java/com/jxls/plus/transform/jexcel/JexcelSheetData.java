package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.SheetData;
import jxl.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Leonid Vysochyn
 */
public class JexcelSheetData extends SheetData {
    Sheet sheet;

    public static JexcelSheetData createSheetData(Sheet sheet){
        JexcelSheetData sheetData = new JexcelSheetData();
        sheetData.sheet = sheet;
        sheetData.sheetName = sheet.getName();
        sheetData.columnWidth = new int[256];
        for(int i = 0; i < 256; i++){
            sheetData.columnWidth[i] = sheet.getColumnView(i).getSize();
        }
        int numberOfRows = sheet.getRows();
        for(int i = 0; i < numberOfRows; i++){
            sheetData.rowDataList.add(JexcelRowData.createRowData(sheetData.sheetName, sheet.getRow(i), sheet.getRowView(i)));
        }
        return sheetData;
    }

    public Sheet getSheet() {
        return sheet;
    }
}
