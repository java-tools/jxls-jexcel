package com.jxls.plus.transform.jexcel;

import com.jxls.plus.common.Context;
import jxl.Cell;
import jxl.Workbook;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.net.URL;

/**
 * Writable cell value implementation for Hyperlink
 * @author Leonid Vysochyn
 *         Date: 6/18/12
 */
public class WritableHyperlinkCell implements WritableCellValue {
    public static final String LINK_URL = "URL";
    public static final String LINK_DOCUMENT = "DOCUMENT";
    public static final String LINK_FILE = "FILE";

    File file;
    URL url;
    String description;

    public WritableHyperlinkCell(URL url, String description) {
        this.url = url;
        this.description = description;
    }

    public WritableHyperlinkCell(File file, String description){
        this.file = file;
        this.description = description;
    }

    public WritableCell writeToCell(WritableSheet sheet, int col, int row, Context context) throws WriteException {
        WritableHyperlink hyperlink = null;
        if( url != null ){
            if( description != null ){
                hyperlink = new WritableHyperlink(col, row, col, row, url, description);
            }else{
                hyperlink = new WritableHyperlink(col, row, url);
            }
        }
        if( file != null ){
            if( description != null ){
                hyperlink = new WritableHyperlink(col, row, file, description);
            }else{
                hyperlink = new WritableHyperlink(col, row, file);
            }
        }
        sheet.addHyperlink(hyperlink);
        Cell cell = sheet.getCell(col, row);
        WritableCell writableCell = new Label(col, row, description != null ? description : url.toString());
        if( cell != null && cell.getCellFormat() != null ){
            writableCell.setCellFormat( cell.getCellFormat() );
        }
        sheet.addCell(writableCell);
        return writableCell;
    }

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    public URL getUrl() {
        return url;
    }

    public void setUrl(URL url) {
        this.url = url;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }
}
