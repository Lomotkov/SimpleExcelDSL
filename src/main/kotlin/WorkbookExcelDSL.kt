package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFWorkbook

fun workbookExcel(block: WorkbookExcelDSL.() -> Unit): XSSFWorkbook =
    WorkbookExcelDSL().apply(block).workbook

class WorkbookExcelDSL {
    val workbook: XSSFWorkbook = XSSFWorkbook()
    val excelStyles by lazy { ExcelStylesDSL(workbook) }
    fun sheet(caption: String, block: SheetExcelDSL.() -> Unit): SheetExcelDSL =
        SheetExcelDSL(caption, workbook.createSheet(caption), this).apply(block)

    fun style(block: CellStyle.() -> Unit): CellStyle =
        workbook.createCellStyle().apply(block)

    fun XSSFCell.setAlignment(alignment: HorizontalAlignment) = CellUtil.setAlignment(this, alignment)

    companion object {
        const val symbolSize = 256
        const val maxChars = 255
        const val defaultColumnSize = 15
    }

}