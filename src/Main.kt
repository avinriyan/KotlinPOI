/**
 * <h1>Read and Write Microsoft Office files</h1>
 * KotlinPOI program implements an application that
 * simply reads and writes Microsoft Office files (docx, xlsx, and pptx)
 * using Apache POI library and written in Kotlin.
 * @see <a href="https://poi.apache.org/apidocs/4.1/">Apache POI</a>
 *
 * @author Avin Riyan
 * @version 1.0
 * @since 2019-12-16
 */
fun main() {
    printPOI()
    //This is an example for Microsoft Windows Operating System
    println("O======================== WRITE ========================O")
    createDocx("C:/Users/Public/Documents/document.docx", "Hello, world!")
    createXlsx("C:/Users/Public/Documents/spreadsheet.xlsx", "Hello, world!")
    createPptx("C:/Users/Public/Documents/presentation.pptx", "Hello, world!")
    println("O======================================================O\n")
    println("O======================== READ ========================O")
    println("Output from reading docx file: "+readDocx("C:/Users/Public/Documents/document.docx"))
    println("Output from reading xlsx file: "+readXlsx("C:/Users/Public/Documents/spreadsheet.xlsx"))
    println("Output from reading pptx file: "+readPptx("C:/Users/Public/Documents/presentation.pptx"))
    println("O======================================================O")
}

/**
 * This method is used to print Apache POI ASCII ART.
 */
fun printPOI(){
    println("     _                     _            ____   ___ ___ \n" +
            "    / \\   _ __   __ _  ___| |__   ___  |  _ \\ / _ \\_ _|\n" +
            "   / _ \\ | '_ \\ / _` |/ __| '_ \\ / _ \\ | |_) | | | | | \n" +
            "  / ___ \\| |_) | (_| | (__| | | |  __/ |  __/| |_| | | \n" +
            " /_/   \\_\\ .__/ \\__,_|\\___|_| |_|\\___| |_|    \\___/___|\n" +
            "         |_| ")
}
