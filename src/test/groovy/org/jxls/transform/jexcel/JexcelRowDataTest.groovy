package org.jxls.transform.jexcel

import jxl.write.Blank
import jxl.write.Formula
import jxl.write.Label
import jxl.write.WritableSheet
import jxl.write.WritableWorkbook
import spock.lang.Specification


/**
 * @author Leonid Vysochyn
 * Date: 2/1/12 2:03 PM
 */
class JexcelRowDataTest extends Specification{
    WritableWorkbook wb;
    BufferedOutputStream outputStream;

    def setup(){
        Locale.setDefault(Locale.ENGLISH)
        outputStream = new BufferedOutputStream(new ByteArrayOutputStream());
        wb = jxl.Workbook.createWorkbook(outputStream)
        WritableSheet sheet = wb.createSheet("sheet 1", 0)
        sheet.addCell(new jxl.write.Number(0, 0, 1.5));
        sheet.addCell(new Label(1, 0, '${x}'))
        sheet.addCell(new Label(2, 0, '${x*y}'))
        sheet.setRowView(0, 123)

        sheet.addCell(new Formula(1, 1, "SUM(A1:A3)"))
        sheet.addCell(new Label(2, 1, '${y*y}'))
        sheet.addCell(new Label(3, 1, '${x} words'))
        sheet.setRowView(1, 456)

        sheet.addCell(new Label(0, 2, "XYZ"))
        sheet.addCell(new Label(1, 2, '${2*y}'))
        sheet.addCell(new Label(2, 2, '${4*4}'))
        sheet.addCell(new Label(3, 2, '${2*x}x and ${2*y}y'))
        sheet.addCell(new Label(4, 2, '${2*x}x and ${2*y} ${cur}'))

        sheet.setColumnView(1, 123)

        WritableSheet sheet2 = wb.createSheet("sheet 2", 1)
        sheet2.addCell(new Blank(0, 0))
    }

    def "test createRowData"(){
        when:
            def rowData = JexcelRowData.createRowData(wb.getSheet(0), 1);
        then:
            rowData.getHeight() == wb.getSheet(0).getRowView(1).getSize()
            rowData.getNumberOfCells() == 4
            rowData.getCellData(2).getCellValue() == '${y*y}'
            rowData.getCellData(2).getSheetName() == "sheet 1"
    }
    
    def "test createRowData for null row"(){
        expect:
            JexcelRowData.createRowData(wb.getSheet(0), 5).getNumberOfCells() == 0
    }
}
