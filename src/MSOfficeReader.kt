import org.apache.poi.xslf.usermodel.XMLSlideShow
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.File
import java.io.FileInputStream

/**
 * This method is used to read simple docx file in a specific location.
 * @param path path where the file will be read
 * Example in Microsoft Windows: "C:/Users/Public/Documents/document.docx"
 * Example in Unix-like OS: "/home/user/Documents/document.docx"
 * @return String This returns text in the file
 */
fun readDocx(path: String): String {
    var text = ""
    val fileInputStream = FileInputStream(File(path))
    val document = XWPFDocument(fileInputStream)
    val paragraphs = document.paragraphs
    for (paragraph in paragraphs) {
        text = paragraph.text
    }
    fileInputStream.close()
    return text
}

/**
 * This method is used to read simple xlsx file in a specific location.
 * @param path path where the file will be read
 * Example in Microsoft Windows: "C:/Users/Public/Documents/workbook.xlsx"
 * Example in Unix-like OS: "/home/user/Documents/workbook.xlsx"
 * @return String This returns text in the file
 */
fun readXlsx(path: String): String {
    var text = ""
    val fileInputStream = FileInputStream(File(path))
    val workbook = XSSFWorkbook(fileInputStream)
    val rows = workbook.getSheetAt(0).getRow(2)
    for (row in rows){
        text = row.stringCellValue
    }
    fileInputStream.close()
    return text
}

/**
 * This method is used to read simple pptx file in a specific location.
 * @param path path where the file will be read
 * Example in Microsoft Windows: "C:/Users/Public/Documents/presentation.pptx"
 * Example in Unix-like OS: "/home/user/Documents/presentation.pptx"
 * @return String This returns text in the file
 */
fun readPptx(path: String): String {
    var text = ""
    val fileInputStream = FileInputStream(File(path))
    val slideShow = XMLSlideShow(fileInputStream)
    val slides = slideShow.slides
    for (slide in slides){
        text = slide.getPlaceholder(0).text
    }
    fileInputStream.close()
    return text
}
