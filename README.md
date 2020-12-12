# KotlinPOI
KotlinPOI program implements an application that simply reads and writes Microsoft Office files (docx, xlsx, and pptx) using [Apache POI](https://poi.apache.org/) library and written in [Kotlin](https://kotlinlang.org/).

## Usage
Function call example that writes docx, xlsx, and pptx files in Windows Operating System
```kotlin
createDocx("C:/Users/Public/Documents/document.docx", "Hello, world!")
createXlsx("C:/Users/Public/Documents/workbook.xlsx", "Hello, world!")
createPptx("C:/Users/Public/Documents/presentation.pptx", "Hello, world!")
```

Function call example that reads docx, xlsx, and pptx files in Windows Operating System
```kotlin
readDocx("C:/Users/Public/Documents/document.docx")
readXlsx("C:/Users/Public/Documents/workbook.xlsx")
readPptx("C:/Users/Public/Documents/presentation.pptx")
```

## License
[Apache License (Version 2.0)](https://www.apache.org/licenses/LICENSE-2.0)
