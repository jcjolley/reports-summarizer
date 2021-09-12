package me.jolley

import org.apache.poi.wp.usermodel.HeaderFooterType
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.tika.Tika
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STLineSpacingRule
import java.io.File
import java.io.FileOutputStream
import java.math.BigInteger
import java.nio.file.Files
import java.nio.file.Paths

fun main(args: Array<String>) {
    val dirRoot = """C:/source-reports/"""
    File(dirRoot)
        .walk()
        .filter { !it.isDirectory }
        .forEach {
            println("File is: ${it.absolutePath}")
            createReport(it)
        }
}

fun createReport(file: File) {
    val tika = Tika();
    val output = tika.parseToString(file)
    val category = getCategory(output)
    val dilution = getSampleDilution(output)
    val internalStandard = getInternalStandard(output)
    val solvent = try {
        getSolvent(output)
    } catch (e: Exception) {
        "Unknown"
    }

    val wordContent = """
    Category:  $category
    GC-MS-FID Method File Name:  $solvent Base Method HP-5 Column
    Sample Dilution:  $dilution in $solvent
    Internal Standard:  $internalStandard
    """.trimIndent()

    createWordDoc(file.name, wordContent)
}

fun createWordDoc(filename:String, contents: String) {
    if (!Paths.get("C:/generated-reports").toFile().exists()) Files.createDirectories(Paths.get("C:/generated-reports"))
    //Blank Document
    val document = XWPFDocument()
    //Write the Document in file system
    val out = FileOutputStream("C:/generated-reports/$filename")
    //create Paragraph
    val header = document.createHeader(HeaderFooterType.DEFAULT)
    val hparagraph = header.createParagraph()
    val hrun = hparagraph.createRun()
    hrun.setText("GC-MS-FID Chromatogram Reference Library -- Univar Solutions")
    val paragraph = document.createParagraph()
    val ppr = paragraph.ctp.pPr ?: paragraph.ctp.addNewPPr()
    val spacing = if (ppr.isSetSpacing) ppr.spacing else ppr.addNewSpacing()
    spacing.lineRule = STLineSpacingRule.AUTO
    spacing.line = BigInteger.valueOf(360)
    val run = paragraph.createRun()
    run.fontFamily = "Calibri"
    run.fontSize = 12

    contents.split('\n').forEach {
        run.setText( it )
        run.addCarriageReturn()
    }
    document.write(out)
    //Close document
    out.close()
}

fun getCategory(sourceReport: String): String {
    val lines = sourceReport.split('\n')
    val categoryIndex = lines.indexOfFirst { """sample description:""".toRegex(RegexOption.IGNORE_CASE).containsMatchIn(it) } + 1
    if (categoryIndex < 1) throw IllegalStateException("Category could not be found, Sample description not found in docx")
    return lines[categoryIndex]
}

fun getGcMsFidTestingMethod(sourceReport: String): String {
    val lines = sourceReport.split('\n')
    val testMethodsIndex = lines.indexOfFirst { """testing methods:""".toRegex(RegexOption.IGNORE_CASE).containsMatchIn(it) }
    if (testMethodsIndex < 0) throw IllegalStateException("Testing methods section could not be found")
    val testMethodLines = lines.drop(testMethodsIndex)
    val gcMsFidSectionIndex = testMethodLines.indexOfFirst { it.contains("GC/MS/FID") }
    if (gcMsFidSectionIndex < 0) throw IllegalStateException("GC/MS/FID Analysis section could not be found")
    return testMethodLines[gcMsFidSectionIndex + 2]
}

fun getSampleDilution(sourceReport: String): String {
    val gcMsFidTestMethod = getGcMsFidTestingMethod(sourceReport)
    val lines = gcMsFidTestMethod.split(". ")
    val dilutionPattern = """(\d+:\d+)""".toRegex()
    val dilutionLine = lines.find { dilutionPattern.containsMatchIn(it) } ?: throw IllegalStateException("Sample dilution could not be found")
    return dilutionPattern.find(dilutionLine)
        ?.groupValues?.get(1)
        ?: throw IllegalStateException("Sample dilution could not be found")
}

fun getSolvent(sourceReport: String): String {
    val possibleSolvents = listOf("methanol", "MeOH",  "tetrahydrofuran", "THF", "Acetone")
    val gcMsFidTestMethodStr = getGcMsFidTestingMethod(sourceReport)
    return possibleSolvents.find { it.toRegex(RegexOption.IGNORE_CASE).containsMatchIn(gcMsFidTestMethodStr)}
        ?.replaceFirstChar { it.titlecaseChar() }
        ?: throw IllegalStateException("No solvent found")
}

fun getInternalStandard(sourceReport: String): String {
    val standardMap = mapOf(
        "dodecane" to "dodecane".toRegex(RegexOption.IGNORE_CASE),
        "1,3 pentanediol" to "pentanediol".toRegex(RegexOption.IGNORE_CASE)
    )
    val gcMsFidTestMethodStr = getGcMsFidTestingMethod(sourceReport)
    return standardMap.entries.find {it.value.containsMatchIn(gcMsFidTestMethodStr)}
        ?.key
        ?.replaceFirstChar { it.titlecaseChar() }
        ?: "N/A"
}