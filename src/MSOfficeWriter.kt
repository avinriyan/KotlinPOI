import org.apache.poi.xslf.usermodel.SlideLayout
import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileOutputStream
import java.nio.file.Paths

/**
 * This method is used to write simple docx file in a specific location.
 * @param path path where the file will be stored
 * Example in Microsoft Windows: "C:/Users/Public/Documents/document.docx"
 * Example in Unix-like OS: "/home/user/Documents/document.docx"
 * @param message message to be written in the file
 * Example: "Hello, world!"
 * @return Nothing.
 */
fun createDocx(path: String, message: String) {
    val paths = Paths.get(path)
    val fileName = paths.fileName
    val fileOutputStream = FileOutputStream(File(path))

    val document = XWPFDocument()
    val paragraph = document.createParagraph()
    val run = paragraph.createRun()
    run.setText(message)

    document.write(fileOutputStream)
    fileOutputStream.close()
    println("$fileName was successfully created and located in $path")
}

/**
 * This method is used to write simple xlsx file in a specific location.
 * @param path path where the file will be stored
 * Example in Microsoft Windows: "C:/Users/Public/Documents/spreadsheet.xlsx"
 * Example in Unix-like OS: "/home/user/Documents/spreadsheet.xlsx"
 * @param message message to be written in the file
 * Example: "Hello, world!"
 * @return Nothing.
 */
fun createXlsx(path: String, message: String) {
    val paths = Paths.get(path)
    val fileName = paths.fileName
    val fileOutputStream = FileOutputStream(File(path))

    val workbook = XSSFWorkbook()
    val sheet = workbook.createSheet("Sheet 1")
    val row = sheet.createRow(2)
    val cell = row.createCell(5)
    cell.setCellValue(message)

    workbook.write(fileOutputStream)
    fileOutputStream.close()
    println("$fileName was successfully created and located in $path")
}

/**
 * This method is used to write simple pptx file in a specific location.
 * @param path path where the file will be stored
 * Example in Microsoft Windows: "C:/Users/Public/Documents/presentation.pptx"
 * Example in Unix-like OS: "/home/user/Documents/presentation.pptx"
 * @param message message to be written in the file
 * Example: "Hello, world!"
 * @return Nothing.
 */
fun createPptx(path: String, message: String) {
    val paths = Paths.get(path)
    val fileName = paths.fileName
    val fileOutputStream = FileOutputStream(File(path))

    val slideShow = XMLSlideShow()
    val slideMaster = slideShow.slideMasters[0]
    val slideLayout = slideMaster.getLayout(SlideLayout.TITLE_ONLY)
    val slide = slideShow.createSlide(slideLayout)
    val title = slide.getPlaceholder(0)
    title.text = message

    slideShow.write(fileOutputStream)
    fileOutputStream.close()
    println("$fileName was successfully created and located in $path")
}