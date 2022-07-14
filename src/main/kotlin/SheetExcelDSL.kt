package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFSheet

class SheetExcelDSL(
    val caption: String,
    private val xssfSheet: XSSFSheet,
    private val workbookExcelDSL: WorkbookExcelDSL
) {
    var style: CellStyle? = null
    private var rowPosition: Int = 0
    private var columnPosition: Int = 0

    fun blank(rowCount: Int = 1) {
        rowPosition += rowCount
    }

    fun mergeRegion(firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) =
        xssfSheet.addMergedRegion(CellRangeAddress(firstRow, lastRow, firstCol, lastCol))

    fun merge(firstRow: Int, firstCol: Int, length: Int, width: Int) =
        xssfSheet.addMergedRegion(CellRangeAddress(firstRow, firstRow + length - 1, firstCol, firstCol + width - 1))

    fun columnSetup(width: Int, position: Int = columnPosition) {
        xssfSheet.setColumnWidth(position, width)
        columnPosition++
    }

    fun columnsSetup(width: Int, count: Int = columnPosition) {
        for (i in 0..count) {
            xssfSheet.setColumnWidth(count, width)
            columnPosition++
        }
    }

    fun columnSetup(valuesList: List<String?>, position: Int = columnPosition) {
        xssfSheet.setColumnWidth(
            position,
            (valuesList.mapNotNull { it?.length }.maxOrNull()
                ?: WorkbookExcelDSL.defaultColumnSize) * WorkbookExcelDSL.symbolSize
        )
        columnPosition++
    }

    fun row(position: Int = rowPosition, block: RowExcelDSL.() -> Unit): RowExcelDSL =
        RowExcelDSL(xssfSheet.createRow(position), style).apply(block).also {
            it.xssfRow.height = it.height
            rowPosition++
        }
}