import com.siams.med.ui.reports.excel.workbookExcel
import java.io.FileOutputStream


fun main() {
    val wb = workbookExcel {
        sheet("test1") {
            style = excelStyles.defaultStyle
            row {
                for (i in 1..9)
                    cell {
                        style = excelStyles.textStyle
                        value = i.toString()
                    }
                cell { }
                cell { }
            }
            row {
                for (i in 1..50)
                    cell {
                        style = excelStyles.textStyle
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
    val out = FileOutputStream("") // your path
    wb.write(out)
}

