package com.jxls.plus.transform.jexcel

import com.jxls.plus.common.Context
import jxl.Hyperlink
import jxl.Workbook
import jxl.write.Blank
import jxl.write.Formula
import jxl.write.Label
import jxl.write.WritableSheet
import jxl.write.WritableWorkbook
import jxl.write.Number

import com.jxls.plus.common.CellRef
import spock.lang.Ignore
import spock.lang.Specification

/**
 * @author Leonid Vysochyn
 * Date: 1/30/12 5:52 PM
 */
class JexcelCellDataTest extends Specification{
    Workbook workbook;
    WritableWorkbook writableWorkbook;

    def setup(){
        def baos = new ByteArrayOutputStream()
        writableWorkbook = Workbook.createWorkbook(baos)
        WritableSheet sheet = writableWorkbook.createSheet("sheet 1", 0)
        sheet.addCell(new Number(0, 0, 1.5));
        sheet.addCell(new Label(1, 0, '${x}'))
        sheet.addCell(new Label(2, 0, '${x*y}'))
        sheet.addCell(new Label(3, 0, '$[B2+B3]'))

        sheet.addCell(new Formula(1, 1, "SUM(A1:A3)"))
        sheet.addCell(new Label(2, 1, '${y*y}'))
        sheet.addCell(new Label(3, 1, '${x} words'))
        sheet.addCell(new Label(4, 1, '$[${myvar}*SUM(A1:A5) + ${myvar2}]'))
        sheet.addCell(new Label(5, 1, '$[SUM(U_(B1,B2)]'))

        sheet.addCell(new Label(0, 2, "XYZ"))
        sheet.addCell(new Label(1, 2, '${2*y}'))
        sheet.addCell(new Label(2, 2, '${4*4}'))
        sheet.addCell(new Label(3, 2, '${2*x}x and ${2*y}y'))
        sheet.addCell(new Label(4, 2, '${2*x}x and ${2*y} ${cur}'))

        WritableSheet sheet2 = writableWorkbook.createSheet("sheet 2", 1)
        sheet2.addCell(new Blank(0, 0))
        sheet2.addCell(new Blank(1, 1))
        sheet2.addCell(new Label(2, 1, '''${util.hyperlink('http://google.com/', 'Google')}'''))
//        sheet2.getRow(1).createCell(2).setCellValue('''${poi.hyperlink('http://google.com/', 'Google', 'URL')}''')
        writableWorkbook.write()
        writableWorkbook.close()
        byte[] workbookBytes = baos.toByteArray()
        workbook = Workbook.getWorkbook(new ByteArrayInputStream(workbookBytes))
    }

    def "test get cell Value"(){
        when:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", row, col), workbook.getSheet(0).getCell(col, row) )
        then:
            assert cellData.getCellValue() == value
        where:
            row | col   | value
            0   | 0     | new Double(1.5)
            0   | 1     | '${x}'
            0   | 2     | '${x*y}'
            1   | 1     | "SUM(A1:A3)"
            2   | 0     | "XYZ"
    }

    def "test evaluate simple expression"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 0, 1), workbook.getSheet(0).getCell(1, 0))
            def context = new Context()
            context.putVar("x", 35)
        expect:
            cellData.evaluate(context) == 35
    }
    
    def "test evaluate multiple regex"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 2, 3), workbook.getSheet(0).getCell(3, 2))
            def context = new Context()
            context.putVar("x", 2)
            context.putVar("y", 3)
        expect:
            cellData.evaluate(context) == "4x and 6y"
    }

    def "test evaluate single expression constant string concatenation"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 1, 3),workbook.getSheet(0).getCell(3,1))
            def context = new Context()
            context.putVar("x", 35)
        expect:
            cellData.evaluate(context) == "35 words"
    }

    def "test evaluate regex with dollar sign"(){
        JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 2, 4), workbook.getSheet(0).getCell(4, 2))
        def context = new Context()
        context.putVar("x", 2)
        context.putVar("y", 3)
        context.putVar("cur", '$')
        expect:
            cellData.evaluate(context) == '4x and 6 $'
    }

    def "test write to another sheet"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 0, 1), writableWorkbook.getSheet(0).getCell(1, 0))
            def context = new Context()
            context.putVar("x", 35)
        when:
            cellData.writeToCell(writableWorkbook.getSheet(1), 0, 0, context)
        then:
            writableWorkbook.getSheet(1).getCell(0, 0).getContents() == "35"
    }

    def "test write user formula"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 0, 3), writableWorkbook.getSheet(0).getCell(3, 0))
            def context = new Context()
        when:
            cellData.writeToCell(writableWorkbook.getSheet(1), 1, 1, context)
        then:
            writableWorkbook.getSheet(1).getCell(1, 1).getContents() == "B2+B3"
    }
    
    def "test write parameterized formula cell"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 1, 4), writableWorkbook.getSheet(0).getCell(4, 1))
            def context = new Context()
            context.putVar("myvar", 2)
            context.putVar("myvar2", 3)
//            writableWorkbook.getSheet(0).createRow(7).createCell(7)
        when:
            cellData.writeToCell(writableWorkbook.getSheet(0), 7, 7, context)
        then:
            writableWorkbook.getSheet(0).getCell(7, 7).getContents() == "2.0*SUM(A1:A5)+3.0"
    }
    
    def "test formula cell check"(){
        when:
            JexcelCellData notFormulaCell = JexcelCellData.createCellData(new CellRef("sheet 1", 0, 1), workbook.getSheet(0).getCell(1, 0))
            JexcelCellData formulaCell1 = JexcelCellData.createCellData(new CellRef("sheet 1", 1, 1), workbook.getSheet(0).getCell(1, 1))
            JexcelCellData formulaCell2 = JexcelCellData.createCellData(new CellRef("sheet 1", 1, 4), workbook.getSheet(0).getCell(4, 1))
            JexcelCellData formulaCell3 = JexcelCellData.createCellData(new CellRef("sheet 1", 0, 3), workbook.getSheet(0).getCell(3, 0))
        then:
            !notFormulaCell.isFormulaCell()
            formulaCell1.isFormulaCell()
            formulaCell2.isFormulaCell()
            formulaCell3.isFormulaCell()
            formulaCell3.getFormula() == "B2+B3"
    }

    def "test write formula with jointed cells"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 1", 1, 5), writableWorkbook.getSheet(0).getCell(5, 1))
            def context = new Context()
        when:
            cellData.writeToCell(writableWorkbook.getSheet(1), 1, 1, context)
        then:
            writableWorkbook.getSheet(1).getCell(1, 1).getContents() == "SUM(U_(B1,B2)"
    }

    // todo:
    @Ignore("Not implemented yet")
    def "test write merged cell"(){

    }

    def "test hyperlink cell"(){
        setup:
            JexcelCellData cellData = JexcelCellData.createCellData(new CellRef("sheet 2", 1, 2), writableWorkbook.getSheet(1).getCell(2, 1))
            def context = new JexcelContext()
      when:
            cellData.writeToCell(writableWorkbook.getSheet(1), 2, 1, context)
      then:
            def sheet = writableWorkbook.getSheet(1)
            Hyperlink[] links = sheet.getHyperlinks()
            links.length == 1
            Hyperlink link = links[0]
            link.isURL()
            link.getURL().toString() == "http://google.com/"
    }

}
