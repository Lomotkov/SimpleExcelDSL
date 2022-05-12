package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

fun workbookExcel(block: WorkbookExcelDSL.() -> Unit): XSSFWorkbook =
    WorkbookExcelDSL().apply(block).workbook

class WorkbookExcelDSL {
    private val sheetList: MutableList<SheetExcelDSL> = mutableListOf()
    val workbook: XSSFWorkbook = XSSFWorkbook()
    fun sheet(caption: String, block: SheetExcelDSL.() -> Unit): SheetExcelDSL =
        SheetExcelDSL(caption, workbook.createSheet(caption)).apply(block)
            .also { addSheet(it) }

//    fun createDocument(): XSSFWorkbook = workbook.apply {
//        sheetList.forEach { sheet ->
//            sheet.xssfSheet.apply {
//                sheet.rows.forEach { (rowNumber, row) ->
//                    row.xssfRow.apply {
//                        row.cells.forEach { (cellNumber, cell) ->
//                            createCell(cellNumber).apply {
//                                setCellValue(cell.value)
//                            }
//                        }
//                    }
//                }
//            }
//        }
//    }

    fun style(block: XSSFCellStyle.() -> Unit): XSSFCellStyle =
        workbook.createCellStyle().apply(block)


    private fun addSheet(sheetExcelDSL: SheetExcelDSL) {
        sheetList.add(sheetExcelDSL)
    }

}

class SheetExcelDSL(
    val caption: String,
    val xssfSheet: XSSFSheet
) {
    var style: XSSFCellStyle? = null
    private var rowPosition: Int = 0
    val rows: MutableMap<Int, RowExcelDSL> = mutableMapOf()

    fun row(position: Int = rowPosition, block: RowExcelDSL.() -> Unit): RowExcelDSL =
        RowExcelDSL(xssfSheet.createRow(position), style).apply(block)
            .also { addRow(it) }

    private fun addRow(rowExcelDSL: RowExcelDSL) {
        rows[rowPosition++] = rowExcelDSL
    }
}

class RowExcelDSL(val xssfRow: XSSFRow, var style: XSSFCellStyle? = null) {
    var cellPosition: Int = 0
    val cells: MutableMap<Int, CellExcelDsl> = mutableMapOf()

    fun cell(value: String = "", block: CellExcelDsl.() -> Unit): CellExcelDsl =
        CellExcelDsl(value, xssfRow.createCell(cellPosition), style)
            .apply(block)
            .also {
                it.changeCell()
                addRow(it)
            }

    private fun addRow(cellExcelDsl: CellExcelDsl) {
        cells[cellPosition++] = cellExcelDsl
    }

}

class CellExcelDsl(var value: String, val xssfCell: XSSFCell, var style: XSSFCellStyle? = null) {
    fun changeCell() {
        xssfCell.setCellValue(value)
        xssfCell.cellStyle = style
    }
}

fun main() {
    val wb = workbookExcel {
        val test1 = style {
            alignment = HorizontalAlignment.CENTER.code
            verticalAlignment = VerticalAlignment.CENTER.code
        }
        val test2 = style {
            borderBottom = BorderStyle.THIN.code
            borderLeft = BorderStyle.THIN.code
            borderRight = BorderStyle.THIN.code
            borderTop = BorderStyle.THIN.code
        }
        val test3 = style {}
        sheet("test1") {
            style = test1
            row {
                style = test2
                for (i in 1..9)
                    cell {
                        style = test1
                        value = i.toString()
                    }
                for (i in 10..20)
                    cell {
                        value = i.toString()
                    }
            }
            row {
                for (i in 1..9)
                    cell {
                        style = test3
                        value = i.toString()
                    }
                cell { }
                cell { }
            }
        }
        sheet("test2") {
            row {
                for (i in 1..9)
                    cell {
                        value = i.toString()
                    }
                cell { }
                cell { }
            }
        }
    }
    val out = FileOutputStream("/home/dima/Downloads/yourFileName.xls")
    wb.write(out)
    println("test")
}