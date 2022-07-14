package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.usermodel.XSSFRow


class RowExcelDSL(
    val xssfRow: XSSFRow,
    var style: CellStyle? = null,
    var height: Short = xssfRow.height
) {
    private var cellPosition: Int = 0

    fun blank(rowCount: Int = 1) {
        cellPosition += rowCount
    }

    fun cell(value: String? = "", block: CellExcelDsl.() -> Unit): CellExcelDsl =
        CellExcelDsl(value, xssfRow.createCell(cellPosition++), style)
            .apply(block)
            .also { it.prepareCell() }
}