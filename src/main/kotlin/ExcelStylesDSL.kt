package com.siams.med.ui.reports.excel

import org.apache.poi.ss.usermodel.*

class ExcelStylesDSL(workbook: Workbook) { //TODO styles what you want
    val defaultStyle: CellStyle = workbook.createCellStyle().apply {
        verticalAlignment = VerticalAlignment.CENTER
        borderTop = BorderStyle.THIN
        borderRight = BorderStyle.THIN
        borderBottom = BorderStyle.THIN
        borderLeft = BorderStyle.THIN
        wrapText = true
    }

    val textStyle: CellStyle = workbook.createCellStyle().apply {
        cloneStyleFrom(defaultStyle)
        dataFormat = workbook.createDataFormat().getFormat("@")
    }

    val headerStyle: CellStyle = workbook.createCellStyle().apply {
        cloneStyleFrom(defaultStyle)
        alignment = HorizontalAlignment.CENTER
        fillForegroundColor = IndexedColors.GREY_25_PERCENT.index
        fillPattern = FillPatternType.SOLID_FOREGROUND
    }
}