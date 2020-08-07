package com.grapecity

import com.grapecity.documents.excel.IWorksheet
import com.grapecity.documents.excel.Workbook
import java.io.File
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.OutputKeys
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult


private const val stringsFileName = "strings.xml"
private const val inputExcelFile = "strings_formatted.xlsx"
private const val outputFolderPath = "output"

private fun generateXml(languageName: String, strings: List<AndroidString>) {
    val documentFactory = DocumentBuilderFactory.newInstance()
    val documentBuilder = documentFactory.newDocumentBuilder()
    val document = documentBuilder.newDocument()

    val resourcesElement = document.createElement("resources")
    document.appendChild(resourcesElement)

    strings.forEach {
        val stringElement = document.createElement("string")
        stringElement.setAttribute("name", it.key)
        if (!it.isTranslatable) {
            stringElement.setAttribute("translatable", it.isTranslatable.toString())
        }
        stringElement.textContent = it.value
        resourcesElement.appendChild(stringElement)
    }

    val transformerFactory = TransformerFactory.newInstance()
    val transformer = transformerFactory.newTransformer()
    transformer.setOutputProperty(OutputKeys.INDENT, "yes");
    val source = DOMSource(document)

    val outputDir = File("$outputFolderPath/$languageName")
    if (!outputDir.exists())
        outputDir.mkdirs()

    val result = StreamResult(File(outputDir, stringsFileName))
    transformer.transform(source, result)
}

private fun isWorksheetValid(worksheet: IWorksheet): Boolean {
    return worksheet.getRange(0, 0).value == "name" &&
            worksheet.getRange(0, 1).value == "translatable" &&
            worksheet.getRange(0, 2).value == "values"
}

fun main() {
    val workbook = Workbook()
    workbook.open(inputExcelFile)
    val worksheet = workbook.worksheets[0]

    if (!isWorksheetValid(worksheet)) {
        throw IllegalStateException("Invalid excel file")
    }

    for (column in 2 until worksheet.usedRange.columnCount) {
        val languageName = worksheet.getRange(0, column).value.toString()
        println("$languageName")
        val androidStrings = arrayListOf<AndroidString>()
        for (row in 1..worksheet.usedRange.rowCount) {
            val value = worksheet.getRange(row, column)?.value?.toString()
            if (!value.isNullOrEmpty()) {
                val key = worksheet.getRange(row, 0).value.toString()
                val isTranslatable = worksheet.getRange(row, 1).value.toString().toBoolean()
                androidStrings.add(AndroidString(key, value, isTranslatable))
            }
        }
        generateXml(languageName, androidStrings)
    }
}