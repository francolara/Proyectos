package com.prestamos.app.ui.screen.export

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
import android.graphics.RectF
import android.graphics.pdf.PdfDocument
import java.io.File
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

fun createDashboardDetalleImage(context: Context, title: String, message: String): File {
    val file = createExportFile(context, "png")
    val bitmap = Bitmap.createBitmap(1200, 1600, Bitmap.Config.ARGB_8888)
    val canvas = Canvas(bitmap)

    canvas.drawColor(Color.WHITE)

    val titlePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 56f
        isFakeBoldText = true
    }
    val bodyPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.DKGRAY
        textSize = 40f
    }

    var y = 140f
    canvas.drawText("Resumen del dashboard", 80f, y, titlePaint)
    y += 90f
    canvas.drawText(title, 80f, y, titlePaint)
    y += 90f

    message.split("\n").forEach { line ->
        canvas.drawText(line, 80f, y, bodyPaint)
        y += 58f
    }

    FileOutputStream(file).use { output ->
        bitmap.compress(Bitmap.CompressFormat.PNG, 100, output)
    }
    return file
}

fun createDashboardDetallePdf(context: Context, title: String, message: String): File {
    val file = createExportFile(context, "pdf")
    val document = PdfDocument()

    val pageWidth = 595
    val pageHeight = 842
    val margin = 28f
    val contentWidth = pageWidth - (margin * 2f)

    val titlePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 20f
        isFakeBoldText = true
    }
    val sectionPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(36, 98, 46)
        textSize = 14f
        isFakeBoldText = true
    }
    val bodyPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.DKGRAY
        textSize = 12f
    }
    val cardPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(239, 247, 239)
    }
    val detailCardPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.rgb(230, 240, 230)
    }

    var pageNumber = 1
    var page = document.startPage(PdfDocument.PageInfo.Builder(pageWidth, pageHeight, pageNumber).create())
    var canvas = page.canvas
    var y = margin + 8f

    fun newPage() {
        document.finishPage(page)
        pageNumber += 1
        page = document.startPage(PdfDocument.PageInfo.Builder(pageWidth, pageHeight, pageNumber).create())
        canvas = page.canvas
        y = margin + 8f
    }

    fun ensureSpace(required: Float) {
        if (y + required > pageHeight - margin) {
            newPage()
        }
    }

    fun wrapText(text: String, paint: Paint, maxWidth: Float): List<String> {
        val clean = text.replace("\t", " ").trim()
        if (clean.isEmpty()) return listOf("")
        val out = mutableListOf<String>()
        var remaining = clean
        while (remaining.isNotEmpty()) {
            val count = paint.breakText(remaining, true, maxWidth, null)
            if (count <= 0) break
            var end = count
            if (end < remaining.length) {
                val chunk = remaining.substring(0, end)
                val lastSpace = chunk.lastIndexOf(' ')
                if (lastSpace > 0) end = lastSpace
            }
            val line = remaining.substring(0, end).trim()
            out += line
            remaining = remaining.substring(end).trimStart()
        }
        return if (out.isEmpty()) listOf(clean) else out
    }

    fun drawWrappedLines(lines: List<String>, paint: Paint, lineHeight: Float) {
        lines.forEach { line ->
            ensureSpace(lineHeight + 2f)
            canvas.drawText(line, margin, y, paint)
            y += lineHeight
        }
    }

    fun drawCard(text: String, paint: Paint, fill: Paint) {
        val wrapped = wrapText(text, paint, contentWidth - 20f)
        val lineHeight = 15f
        val cardHeight = (wrapped.size * lineHeight) + 14f
        ensureSpace(cardHeight + 6f)
        val rect = RectF(margin, y - 12f, margin + contentWidth, y - 12f + cardHeight)
        canvas.drawRoundRect(rect, 8f, 8f, fill)
        var lineY = y
        wrapped.forEach { line ->
            canvas.drawText(line, margin + 10f, lineY, paint)
            lineY += lineHeight
        }
        y = rect.bottom + 10f
    }

    drawWrappedLines(listOf("Resumen del dashboard"), titlePaint, 24f)
    drawWrappedLines(wrapText(title, titlePaint, contentWidth), titlePaint, 24f)
    y += 6f

    val lines = message.lines()
    lines.forEach { raw ->
        val line = raw.trim()
        when {
            line.isBlank() -> y += 6f
            line.equals("Resumen", ignoreCase = true) ||
                line.equals("\uD83D\uDCCC Resumen", ignoreCase = true) ||
                line.contains("Detalle de prestamos", ignoreCase = true) ||
                line.equals("Detalle", ignoreCase = true) ||
                line.equals("\uD83D\uDCCB Detalle", ignoreCase = true) ||
                line.equals("Total cobrado", ignoreCase = true) ||
                line.equals("Total ganado", ignoreCase = true) ||
                line.equals("Cronograma de cuotas", ignoreCase = true) ||
                line.equals("\uD83D\uDCCB Cronograma de cuotas", ignoreCase = true) -> {
                ensureSpace(22f)
                canvas.drawText(line, margin, y, sectionPaint)
                y += 18f
            }
            line.startsWith("- ") ||
                line.contains("|") ||
                line.startsWith("\uD83D\uDC64") ||
                line.startsWith("\uD83D\uDCC4") ||
                line.startsWith("\uD83D\uDCC5") ||
                line.startsWith("\uD83D\uDCB0") ||
                line.startsWith("\uD83D\uDCCC") -> {
                drawCard(line, bodyPaint, detailCardPaint)
            }
            line.contains(":") -> {
                drawCard(line, bodyPaint, cardPaint)
            }
            else -> drawWrappedLines(wrapText(line, bodyPaint, contentWidth), bodyPaint, 16f)
        }
    }

    document.finishPage(page)
    FileOutputStream(file).use { output ->
        document.writeTo(output)
    }
    document.close()
    return file
}

private fun createExportFile(context: Context, extension: String): File {
    val formatter = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault())
    val fileName = "dashboard_detalle_${formatter.format(Date())}.$extension"
    val exportDir = File(context.cacheDir, "exports").apply { mkdirs() }
    return File(exportDir, fileName)
}
