package com.jxls.plus.transform.jexcel

import jxl.CellType
import jxl.Workbook
import jxl.format.CellFormat
import jxl.format.Colour
import jxl.format.Font
import jxl.format.Format
import jxl.format.PaperSize
import jxl.write.Formula
import jxl.write.Label
import jxl.write.WritableCellFormat
import jxl.write.WritableFont
import jxl.write.WritableSheet
import jxl.write.WritableWorkbook
import jxl.write.Number
import spock.lang.Specification


import com.jxls.plus.common.Context
import com.jxls.plus.common.CellData
import com.jxls.plus.common.CellRef
import com.jxls.plus.common.AreaRef
import com.jxls.plus.common.ImageType

/**
 * @author Leonid Vysochyn
 * Date: 1/23/12 3:23 PM
 */
class JexcelTransformerTest extends Specification{
    byte[] workbookBytes;

    WritableCellFormat customStyle;

    def setup(){
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(outputStream)
        WritableFont font = new WritableFont(WritableFont.COURIER, 24, WritableFont.NO_BOLD, true)
        font.setColour(Colour.ORANGE)
        customStyle = new WritableCellFormat(font)
        customStyle.setBackground(Colour.AQUA)
        WritableSheet sheet = writableWorkbook.createSheet("sheet 1", 0)
        sheet.getSettings().setBottomMargin(15)
        sheet.getSettings().setPaperSize( PaperSize.A3 )
        sheet.addCell(new Number(0, 0, 1.5))
        sheet.addCell(new Label(1, 0, '${x}', customStyle))
        def label = new Label(2, 0, '${x*y}')
        JexcelUtil.setCellComment(label, "comment 1")
        sheet.addCell(label)
        sheet.addCell(new Label(3, 0, 'Merged value'))
        sheet.mergeCells(3, 0, 4, 1)
        sheet.setRowView(0, 23)
        sheet.setColumnView(1, 123)
        def formulaCell = new Formula(1, 1, "SUM(A1:A3)")
        JexcelUtil.setCellComment( formulaCell, "comment 2")
        sheet.addCell(formulaCell)
        sheet.addCell(new Label(2, 1, '${y*y}'))
        sheet.setRowView(1, 456)
        sheet.addCell(new Label(0, 2, "XYZ"))
        sheet.addCell(new Label(1, 2, '${2*y}'))
        sheet.addCell(new Label(2, 2, '${4*4}'))
        sheet.addCell(new Label(3, 2, '${2*x}x and ${2*y}y'))
        sheet.addCell(new Label(4, 2, '$[${myvar}*SUM(A1:A5) + ${myvar2}]'))
        writableWorkbook.write()
        writableWorkbook.close()
        workbookBytes = outputStream.toByteArray()
    }

    def "test template cells storage"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
        then:
            CellData cellData = jexcelTransformer.getCellData(new CellRef("sheet 1", 0, 1))
            cellData instanceof JexcelCellData
            cellFormatEquals(((JexcelCellData)cellData).getCellFormat(), customStyle)
            //((JexcelCellData)cellData).getCellFormat() == customStyle
            assert jexcelTransformer.getCellData(new CellRef("sheet 1", row, col)).getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
            2   | 4     |  '$[${myvar}*SUM(A1:A5) + ${myvar2}]'
    }

    private static boolean cellFormatEquals(CellFormat format1, CellFormat format2){
        if ( format1 == null && format2 == null ) {
            return true
        }
        if ( format1 != null && format2 != null ){
            return format1.alignment == format2.alignment && format1.backgroundColour == format2.backgroundColour &&
                    formatEquals(format1.format, format2.format) && format1.indentation == format2.indentation &&
                    format1.orientation == format2.orientation && format1.verticalAlignment == format2.verticalAlignment &&
                    format1.shrinkToFit == format2.shrinkToFit && format1.wrap == format2.wrap && fontEquals(format1.font, format2.font);
        }
        return false
    }

    private static boolean formatEquals(Format format1, Format format2){
        if ( format1 == null && format2 == null ) return true;
        if ( format1 == null ){
            return format2.formatString.length() == 0
        }
        if ( format2 == null ){
            return format1.formatString.length() == 0
        }
        return format1.formatString == format2.formatString
    }

    private static boolean fontEquals(Font font1, Font font2){
        if ( font1 == null && font2 == null ) return true;
        if ( font1 != null && font2 != null ){
            return font1.boldWeight == font2.boldWeight && font1.colour == font2.colour && font1.italic == font2.italic &&
                    font1.name == font2.name && font1.pointSize == font2.pointSize && font1.struckout == font2.struckout &&
                    font1.underlineStyle == font2.underlineStyle && font1.scriptStyle == font2.scriptStyle
        }
        return false;
    }

    def "test transform string var"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            context.putVar("x", "Abcde")
            def writableWorkbook = jexcelTransformer.getWritableWorkbook()
            WritableSheet sheet = writableWorkbook.getSheet(0)
        when:
            jexcelTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 1", 7, 7), context)
        then:

            sheet.getCell(7, 7).contents == "Abcde"
            sheet.getColumnView(7).size == 123 *256
            sheet.getRowView(7).size == 23
            cellFormatEquals(sheet.getCell(7, 7).cellFormat, customStyle)
    }

    def "test transform numeric var"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            context.putVar("x", 3)
            context.putVar("y", 5)
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",0, 2), new CellRef("sheet 2",7, 7), context)
        then:
            WritableSheet sheet = jexcelTransformer.getWritableWorkbook().getSheet("sheet 2")
            sheet.getCell(7, 7).getType() == CellType.NUMBER
            ((Number)sheet.getCell(7, 7)).value == 15
    }

    def "test transform formula cell"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",1, 1), new CellRef("sheet 2",7, 7), context)
        then:
            WritableSheet sheet = jexcelTransformer.getWritableWorkbook().getSheet("sheet 2")
            sheet.getCell(7, 7).type == CellType.ERROR
             sheet.getCell(7, 7) instanceof jxl.write.Formula
            sheet.getCell(7, 7).contents == "SUM(A1:A3)"
    }

    def "test transform a cell to other sheet"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",0, 1), new CellRef("sheet2", 7, 7), context)
        then:
//            WritableSheet sheet = jexcelTransformer.getWritableWorkbook().getSheet("sheet 1")
            WritableSheet sheet1 = jexcelTransformer.getWritableWorkbook().getSheet("sheet2")
            sheet1.getCell(7, 7).contents == "Abcde"
            sheet1.getSettings().bottomMargin == 15
            sheet1.getSettings().paperSize == PaperSize.A3
    }
    
    def "test transform multiple times"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context1 = new Context()
            context1.putVar("x", "Abcde")
            def context2 = new Context()
            context2.putVar("x", "Fghij")
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",0, 1), new CellRef("sheet 1",5, 1), context1)
            jexcelTransformer.transform(new CellRef("sheet 1",0, 1), new CellRef("sheet 2",10, 1), context2)
        then:
            def writableWorkbook = jexcelTransformer.getWritableWorkbook()
            WritableSheet sheet = writableWorkbook.getSheet("sheet 1")
            WritableSheet sheet2 = writableWorkbook.getSheet("sheet 2")
            sheet.getCell(1, 5).contents == "Abcde"
            sheet2.getCell(1, 10).contents == "Fghij"
    }

    def "test transform overridden cells"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context1 = new Context()
            context1.putVar("x", "Abcde")
            def context2 = new Context()
            context2.putVar("x", "Fghij")
        when:
            jexcelTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 1", 5, 1), context1)
            jexcelTransformer.transform(new CellRef("sheet 1", 0, 0), new CellRef("sheet 2", 0, 1), context1)
            jexcelTransformer.transform(new CellRef("sheet 1", 0, 1), new CellRef("sheet 2", 10, 1), context2)
        then:
            def writableWorkbook = jexcelTransformer.getWritableWorkbook()
            WritableSheet sheet = writableWorkbook.getSheet("sheet 1")
            sheet.getCell(1, 5).contents == "Abcde"
            WritableSheet sheet2 = writableWorkbook.getSheet("sheet 2")
            sheet2.getCell(1, 0).contents == sheet.getCell(0,0).contents
            sheet2.getCell(1, 10).type == CellType.LABEL
            sheet2.getCell(1, 10).contents == "Fghij"
    }

    def "test multiple expressions in a cell"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",2, 3), new CellRef("sheet 2",7, 7), context)
        then:
            WritableSheet sheet = jexcelTransformer.getWritableWorkbook().getSheet("sheet 2")
            sheet.getCell(7, 7).contents == "4x and 6y"
    }

    def "test ignore source column and row props"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            jexcelTransformer.setIgnoreColumnProps(true)
            jexcelTransformer.setIgnoreRowProps(true)
            def context = new Context()
            context.putVar("x", "Abcde")
        when:
            jexcelTransformer.transform(new CellRef("sheet 1",0, 1), new CellRef("sheet 2",7, 7), context)
        then:
            def writableWorkbook = jexcelTransformer.getWritableWorkbook()
            WritableSheet sheet1 = writableWorkbook.getSheet("sheet 1")
            WritableSheet sheet2 = writableWorkbook.getSheet("sheet 2")
            sheet2.getColumnView(7).size != sheet1.getColumnView(1).size
            sheet1.getRowView(0).size != sheet2.getRowView(7).size
    }

    def "test set formula value"(){
        given:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
        when:
            jexcelTransformer.setFormula(new CellRef("sheet 2",1, 1), "SUM(B1:B5)")
        then:
            jexcelTransformer.getWritableWorkbook().getSheet("sheet 2").getCell(1, 1).contents == "SUM(B1:B5)"
    }

    def "test get formula cells"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            context.putVar("myvar", 2)
            context.putVar("myvar2", 4)
            jexcelTransformer.transform(new CellRef("sheet 1",2,4), new CellRef("sheet 2", 10,10), context)
            def formulaCells = jexcelTransformer.getFormulaCells()
        then:
            formulaCells.size() == 2
            formulaCells.contains(new CellData("sheet 1",1,1, CellData.CellType.FORMULA, "SUM(A1:A3)"))
            formulaCells.contains(new CellData("sheet 1",2,4, CellData.CellType.STRING, '$[${myvar}*SUM(A1:A5) + ${myvar2}]'))
    }

    def "test get target cells"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 2",10,10), context)
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 1",10,12), context)
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 1",10,14), context)
            jexcelTransformer.transform(new CellRef("sheet 1",2,1), new CellRef("sheet 2",20,11), context)
            jexcelTransformer.transform(new CellRef("sheet 1",2,1), new CellRef("sheet 1",20,12), context)
        then:
            jexcelTransformer.getTargetCellRef(new CellRef("sheet 1",1,1)).toArray() == [new CellRef("sheet 2",10,10), new CellRef("sheet 1",10,12), new CellRef("sheet 1",10,14)]
            jexcelTransformer.getTargetCellRef(new CellRef("sheet 1",2,1)).toArray() == [new CellRef("sheet 2",20,11), new CellRef("sheet 1",20,12)]
    }

    def "test reset target cells"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 1",10,10), context)
            jexcelTransformer.transform(new CellRef("sheet 1",2,1), new CellRef("sheet 1",20,11), context)
            jexcelTransformer.resetTargetCellRefs()
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 2",10,12), context)
            jexcelTransformer.transform(new CellRef("sheet 1",1,1), new CellRef("sheet 1",10,14), context)
            jexcelTransformer.transform(new CellRef("sheet 1",2,1), new CellRef("sheet 1",20,12), context)
        then:
            jexcelTransformer.getTargetCellRef(new CellRef("sheet 1",1,1)).toArray() == [new CellRef("sheet 2",10,12), new CellRef("sheet 1",10,14)]
            jexcelTransformer.getTargetCellRef(new CellRef("sheet 1",2,1)).toArray() == [new CellRef("sheet 1",20,12)]
    }

    def "test transform merged cells"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def context = new Context()
            jexcelTransformer.getWritableWorkbook().getSheet(0).getMergedCells().length == 1
            jexcelTransformer.transform(new CellRef("sheet 1",0,3), new CellRef("sheet 1",10,10), context)
        then:
            WritableSheet sheet = jexcelTransformer.getWritableWorkbook().getSheet("sheet 1")
            sheet.getMergedCells().length == 2
            jxl.Range range = sheet.getMergedCells()[1]
            range.getTopLeft().column == 10
            range.getTopLeft().row == 10
            range.getBottomRight().column == 11
            range.getBottomRight().row == 11
    }

    def "test clear cell"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            jexcelTransformer.clearCell(new CellRef("'sheet 1'!B1"))
        then:
            def writableWorkbook = jexcelTransformer.getWritableWorkbook()
            def cell = writableWorkbook.getSheet(0).getCell(1, 0)
            cell.type == CellType.EMPTY
            cell.contents == ""
            cell.cellFormat != customStyle
//            cell.cellFormat == writableWorkbook.getCellStyleAt((short)0)
    }

    def "test get commented cells"(){
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            def commentedCells = jexcelTransformer.getCommentedCells()
        then:
            commentedCells.size() == 2
            commentedCells.get(0).getCellComment() == "comment 1"
            commentedCells.get(1).getCellComment() == "comment 2"
    }

    //@Ignore("the test does not work with dynamically created workbook for some reason")
    def "test addImage"(){
        given:
            InputStream imageInputStream = JexcelTransformerTest.class.getResourceAsStream("ja.png");
            byte[] imageBytes = JexcelUtil.imageStreamToByteArray(imageInputStream);
        when:
            InputStream inputStream = new BufferedInputStream(new ByteArrayInputStream(workbookBytes))
            OutputStream outputStream = new BufferedOutputStream(new ByteArrayOutputStream())
            def jexcelTransformer = JexcelTransformer.createTransformer(inputStream, outputStream)
            jexcelTransformer.addImage(new AreaRef("'sheet 1'!A1:C10"), imageBytes, ImageType.PNG);
        then:
            jexcelTransformer.getWritableWorkbook().getSheet('sheet 1').getNumberOfImages() == 1
    }

}
