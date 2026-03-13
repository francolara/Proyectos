package com.prestamos.app.ui.screen.export

import android.content.Context
import android.graphics.Bitmap
import android.graphics.Canvas
import android.graphics.Color
import android.graphics.Paint
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
    val pageInfo = PdfDocument.PageInfo.Builder(595, 842, 1).create()
    val page = document.startPage(pageInfo)

    val titlePaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.BLACK
        textSize = 18f
        isFakeBoldText = true
    }
    val bodyPaint = Paint(Paint.ANTI_ALIAS_FLAG).apply {
        color = Color.DKGRAY
        textSize = 14f
    }

    val canvas = page.canvas
    var y = 60f
    canvas.drawText("Resumen del dashboard", 40f, y, titlePaint)
    y += 30f
    canvas.drawText(title, 40f, y, titlePaint)
    y += 30f

    message.split("\n").forEach { line ->
        canvas.drawText(line, 40f, y, bodyPaint)
        y += 24f
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
