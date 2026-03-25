package com.prestamos.app.ui.screen.export

import android.content.Context
import android.graphics.Color
import android.graphics.Paint
import android.graphics.pdf.PdfDocument
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

enum class TableAlign { LEFT, CENTER, RIGHT }

data class ReportTableColumn(
    val title: String,
    val weight: Float = 1f,
    val align: TableAlign = TableAlign.LEFT
)

data class ReportTable(
    val columns: List<ReportTableColumn>,
    val rows: List<List<String>>
)

data class ReportPdfSection(
    val title: String,
    val table: ReportTable
)

data class ReportPdfPayload(
    val appName: String,
    val reportType: String,
    val filter: String,
    val generatedAt: String,
    val sections: List<ReportPdfSection>
)

fun createReportesPdf(context: Context, payload: ReportPdfPayload): File {
    val file = createReportesExportFile(context)
    val document = PdfDocument()

    val pageWidth = 842
    val pageHeight = 595
    val margin = 24f
    val usableWidth = pageWidth - (margin * 2f)

    val titlePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 14f
        isFakeBoldText = true
    }
    val metaPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.DKGRAY
        textSize = 9f
    }
    val sectionPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(36, 98, 46)
        textSize = 10f
        isFakeBoldText = true
    }
    val headerBgPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(220, 235, 246)
    }
    val linePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(160, 160, 160)
        strokeWidth = 1f
    }
    val cellPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 8f
    }
    val headerTextPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 8f
        isFakeBoldText = true
    }

    var pageNumber = 1
    var page = document.startPage(PdfDocument.PageInfo.Builder(pageWidth, pageHeight, pageNumber).create())
    var canvas = page.canvas
    var y = margin

    fun startNewPage() {
        document.finishPage(page)
        pageNumber += 1
        page = document.startPage(PdfDocument.PageInfo.Builder(pageWidth, pageHeight, pageNumber).create())
        canvas = page.canvas
        y = margin
    }

    fun ensureSpace(required: Float) {
        if (y + required > pageHeight - margin) {
            startNewPage()
        }
    }

    fun fitText(text: String, paint: Paint, maxWidth: Float): String {
        val clean = text.replace("\n", " ").replace("\r", " ").trim()
        if (clean.isEmpty()) return "-"
        if (paint.measureText(clean) <= maxWidth) return clean
        val ellipsis = "..."
        var result = clean
        while (result.isNotEmpty() && paint.measureText(result + ellipsis) > maxWidth) {
            result = result.dropLast(1)
        }
        return if (result.isBlank()) ellipsis else result + ellipsis
    }

    fun drawTable(table: ReportTable) {
        if (table.columns.isEmpty()) return
        val totalWeight = table.columns.sumOf { it.weight.toDouble() }.toFloat().coerceAtLeast(1f)
        val colWidths = table.columns.map { (it.weight / totalWeight) * usableWidth }
        val headerHeight = 18f
        val rowHeight = 16f

        ensureSpace(headerHeight + rowHeight)

        var x = margin
        table.columns.forEachIndexed { index, col ->
            val w = colWidths[index]
            canvas.drawRect(x, y, x + w, y + headerHeight, headerBgPaint)
            canvas.drawRect(x, y, x + w, y + headerHeight, linePaint)
            val txt = fitText(col.title, headerTextPaint, w - 6f)
            canvas.drawText(txt, x + 3f, y + 12f, headerTextPaint)
            x += w
        }
        y += headerHeight

        table.rows.forEach { row ->
            ensureSpace(rowHeight)
            var colX = margin
            table.columns.forEachIndexed { index, col ->
                val w = colWidths[index]
                canvas.drawRect(colX, y, colX + w, y + rowHeight, linePaint)
                val value = row.getOrElse(index) { "-" }
                val txt = fitText(value, cellPaint, w - 6f)
                val textX = when (col.align) {
                    TableAlign.LEFT -> colX + 3f
                    TableAlign.CENTER -> colX + (w / 2f) - (cellPaint.measureText(txt) / 2f)
                    TableAlign.RIGHT -> colX + w - cellPaint.measureText(txt) - 3f
                }
                canvas.drawText(txt, textX, y + 11f, cellPaint)
                colX += w
            }
            y += rowHeight
        }
    }

    ensureSpace(56f)
    canvas.drawText(payload.appName, margin, y + 12f, titlePaint)
    y += 16f
    canvas.drawText("Tipo de reporte: ${payload.reportType}", margin, y + 8f, metaPaint)
    y += 12f
    canvas.drawText("Filtro: ${payload.filter}", margin, y + 8f, metaPaint)
    y += 12f
    canvas.drawText("Generado: ${payload.generatedAt}", margin, y + 8f, metaPaint)
    y += 18f

    payload.sections.forEachIndexed { index, section ->
        ensureSpace(20f)
        canvas.drawText(section.title, margin, y + 8f, sectionPaint)
        y += 12f
        drawTable(section.table)
        y += if (index < payload.sections.lastIndex) 12f else 0f
    }

    document.finishPage(page)
    FileOutputStream(file).use { output ->
        document.writeTo(output)
    }
    document.close()
    return file
}

private fun createReportesExportFile(context: Context): File {
    val formatter = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault())
    val fileName = "reporte_clientes_${formatter.format(Date())}.pdf"
    val exportDir = File(context.cacheDir, "exports").apply { mkdirs() }
    return File(exportDir, fileName)
}
