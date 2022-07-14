package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont

class CellExcelDsl(
    var value: String?,
    private val xssfCell: XSSFCell,
    var style: CellStyle? = null
) {
    internal fun prepareCell() {
        if (value != "") xssfCell.setCellValue(value)
        xssfCell.cellStyle = style as XSSFCellStyle?
    }


    operator fun String.unaryPlus() {
        xssfCell.setCellValue(this)
    }

    fun font(init: XSSFFont.() -> Unit) =
        CellUtil.setFont(xssfCell, xssfCell.sheet.workbook.createFont().apply { init() })

}