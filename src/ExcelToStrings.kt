package com.grapecity

import com.grapecity.documents.excel.IWorksheet
import com.grapecity.documents.excel.Workbook
import java.io.File
import java.util.regex.Pattern
import javax.xml.parsers.DocumentBuilderFactory
import javax.xml.transform.OutputKeys
import javax.xml.transform.TransformerFactory
import javax.xml.transform.dom.DOMSource
import javax.xml.transform.stream.StreamResult


private const val stringsFileName = "strings.xml"
private const val inputExcelFile = "2021-05-12-App-lang-android-may.xlsx"
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
    transformer.setOutputProperty(OutputKeys.INDENT, "yes")
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


private val formatters = listOf("%s")
private val errorsMap = mutableMapOf<String, Int>()

//Appear at least twice in one strings (open and close)
private val customFormatters = listOf("[\$click]")

private fun String.calculateFormattersCount(): Int {
    return formatters.sumBy { countMatches(it) }
}

private fun String.calculateCustomFormatters(): Int {
    return customFormatters.sumBy { countMatches(it) }
}

private var lastCheckedLanguage = ""

private fun checkStringValid(originalStrings: List<AndroidString>, translate: AndroidString, language: String) {
    val originalString = originalStrings.find { it.key == translate.key }?.value

    checkNotNull(originalString) {
        println("Warning: String with key ${translate.key} ($language) was not found in original strings list")
        return
    }

    val originalFormattersCount = originalString.calculateFormattersCount()
    val translateFormattersCount = translate.value.calculateFormattersCount()
    if (originalFormattersCount != translateFormattersCount && originalFormattersCount != 0) { // originalFormattersCount != 0 as we have some strings that include %s only in translated value for -de language
        if(lastCheckedLanguage!=language){
            lastCheckedLanguage = language
            println()
            println(language)
        }
        println("Error! String with key ${translate.key} ($language) has wrong count of formatters (${formatters[0]}). Original count: $originalFormattersCount, Actual count: $translateFormattersCount. \nOriginal string: $originalString. \nTranslated string: ${translate.value}\n")
        val currentErrorsCount = errorsMap[language]?:0
        errorsMap[language] = currentErrorsCount+1
    }
    val originalCustomFormattersCount = originalString.calculateCustomFormatters()
    val translateCustomFormattersCount = translate.value.calculateCustomFormatters()
    if (originalCustomFormattersCount != translateCustomFormattersCount) {
        if(lastCheckedLanguage!=language){
            lastCheckedLanguage = language
            println()
            println(language)
        }
        println("Error! String with key ${translate.key} ($language) has wrong count of custom formatters (${customFormatters[0]}). Original count: $originalCustomFormattersCount, Actual count: $translateCustomFormattersCount. \nOriginal string: $originalString. \nTranslated string: ${translate.value}\n")
        val currentErrorsCount = errorsMap[language]?:0
        errorsMap[language] = currentErrorsCount+1
    }
}

private fun checkStringsValid(
    originalStrings: List<AndroidString>,
    translatedStrings: List<AndroidString>,
    language: String
) {
    translatedStrings.forEach { checkStringValid(originalStrings, it, language) }
}


private fun checkCustomFormattersCount(originalStrings: List<AndroidString>, language: String) {
    originalStrings.forEach {
        val customFormattersCount = it.value.calculateCustomFormatters()
        if (customFormattersCount % 2 != 0) {
            println("Error! String from original file ${it.key} ($language) has wrong count of custom formatters ($customFormattersCount). It has to have opening tag and closable tag. ${it.value}")
        }
    }
}

private fun String.countMatches(pattern: String): Int {
    //Add spaces before and after for case when pattern is in end or in start
    return " $this ".split(pattern)
        .dropLastWhile { it.isEmpty() }
        .toTypedArray().size - 1
}

fun main() {
    val workbook = Workbook()
    workbook.open(inputExcelFile)
    val worksheet = workbook.worksheets[0]

    if (!isWorksheetValid(worksheet)) {
        throw IllegalStateException("Invalid excel file")
    }

    // Strings from default language (values.xml)
    var originalStrings = arrayListOf<AndroidString>()

    for (column in 2 until worksheet.usedRange.columnCount) {
        val languageName = worksheet.getRange(0, column).value.toString()
//        println(languageName)
        val androidStrings = arrayListOf<AndroidString>()
//        println("$column ${worksheet.usedRange.rowCount}")
        for (row in 1..worksheet.usedRange.rowCount) {
            val value = worksheet.getRange(row, column)?.value?.toString()
            if (!value.isNullOrEmpty()) {
                val key = worksheet.getRange(row, 0).value.toString()
                val isTranslatable = worksheet.getRange(row, 1).value.toString().toBoolean()
                androidStrings.add(AndroidString(key, value, isTranslatable))
            }
        }
        if (languageName == "values") {
            originalStrings = androidStrings
            checkCustomFormattersCount(originalStrings, languageName)
        } else {
            checkStringsValid(originalStrings, androidStrings, languageName)
        }
        generateXml(languageName, androidStrings)
    }
    println("Total errors: ${errorsMap.values.sum()}")
    println("Errors by language: $errorsMap")
}