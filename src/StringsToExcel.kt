package com.grapecity

import com.grapecity.documents.excel.IWorksheet
import com.grapecity.documents.excel.Workbook
import org.w3c.dom.Document
import org.w3c.dom.Element
import java.io.File
import javax.xml.parsers.DocumentBuilderFactory


private const val stringsFileName = "strings.xml"
private const val resourcesFolderPath = "res"

private fun getDefaultValuesFile(): File {
    val resFolder = File(resourcesFolderPath)
    val defaultValuesFolder = File(resFolder, "values")
    return File(defaultValuesFolder, stringsFileName)
}

private fun getAllValuesFiles(): List<File>? {
    val resFolder = File(resourcesFolderPath)
    return resFolder.listFiles { file ->
        file.isDirectory && file.name.startsWith("values-") && File(file, stringsFileName).exists()
    }?.map {
        File(it, stringsFileName)
    }?.toList()
}

private fun getAndroidStringsFromFile(file: File): List<AndroidString> {
    val dbFactory = DocumentBuilderFactory.newInstance()
    val dBuilder = dbFactory.newDocumentBuilder()
    val doc: Document = dBuilder.parse(file)
    doc.documentElement.normalize()
    val stringNodes = doc.getElementsByTagName("string")

    val androidStrings = arrayListOf<AndroidString>()
    for (i in 0 until stringNodes.length) {
        val stringElement = stringNodes.item(i) as Element
        val keyName = stringElement.getAttribute("name")

        val isTranslatable = if (stringElement.hasAttribute("translatable"))
            stringElement.getAttribute("translatable")?.toBoolean() ?: true
        else true

        val keyValue = stringElement.textContent
        androidStrings.add(AndroidString(keyName, keyValue, isTranslatable))
    }
    println("Language: ${file.parentFile.name}; Total strings: ${androidStrings.size}; Untranslatable: ${androidStrings.count { !it.isTranslatable }}; Total strings length: ${androidStrings.sumBy { it.value.length }} ")
    return androidStrings
}

private fun writeDownDefaultStrings(worksheet: IWorksheet) {
    val defaultFile = getDefaultValuesFile()
    val defaultValuesStrings = getAndroidStringsFromFile(defaultFile)

    val defaultFileValues = arrayListOf(
        arrayOf("name", "translatable", defaultFile.parentFile.name)
    )
    defaultValuesStrings.forEach {
        defaultFileValues.add(arrayOf(it.key, it.isTranslatable.toString(), it.value))
    }

    val range = worksheet.getRange(0, 0, defaultFileValues.size, defaultFileValues[0].size)
    range.value = defaultFileValues.toTypedArray()
}

/**
 * @return map of already filled keys and its rows
 */
private fun getKeyRows(worksheet: IWorksheet): Map<String, Int> {
    val keyRows = hashMapOf<String, Int>()
    for (row in 1..worksheet.usedRange.rowCount) {
        val value = worksheet.getRange(row, 0)?.value?.toString()
        if (!value.isNullOrBlank())
            keyRows[value] = row
    }
    return keyRows
}

private fun writeDownTranslations(worksheet: IWorksheet) {
    var column = worksheet.usedRange.columnCount //start column for translations
    val keyRows = getKeyRows(worksheet)
    getAllValuesFiles()?.forEach { file ->
        val androidStrings = getAndroidStringsFromFile(file)
        worksheet.getRange(0, column).value = file.parentFile.name
        androidStrings.forEach {
            val row = keyRows[it.key]
            row?.let { r ->
                worksheet.getRange(r, column).value = it.value
            } ?: println("Warning: No default value found for key ${it.key} for language ${file.parentFile.name}")
        }
        column++
    }
}

fun main() {
    val workbook = Workbook()
    val worksheet = workbook.worksheets.get(0)

    writeDownDefaultStrings(worksheet)
    writeDownTranslations(worksheet)
    workbook.save("strings.xlsx")
}