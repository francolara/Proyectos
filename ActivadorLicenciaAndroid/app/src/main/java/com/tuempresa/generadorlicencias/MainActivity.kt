package com.tuempresa.generadorlicencias

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.os.Bundle
import android.widget.ArrayAdapter
import android.widget.Button
import android.widget.EditText
import android.widget.Spinner
import android.widget.TextView
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import java.security.MessageDigest
import java.util.Locale

class MainActivity : AppCompatActivity() {

    private lateinit var etDeviceCode: EditText
    private lateinit var spinnerTipo: Spinner
    private lateinit var tvResultado: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        etDeviceCode = findViewById(R.id.etDeviceCode)
        spinnerTipo = findViewById(R.id.spinnerTipo)
        tvResultado = findViewById(R.id.tvResultado)

        configurarSpinner()

        findViewById<Button>(R.id.btnGenerar).setOnClickListener {
            generarClaveDesdeUI()
        }

        findViewById<Button>(R.id.btnCopiar).setOnClickListener {
            copiarResultado()
        }

        findViewById<Button>(R.id.btnLimpiar).setOnClickListener {
            limpiarCampos()
        }
    }

    private fun configurarSpinner() {
        val tipos = listOf("ANUAL", "FULL")
        val adapter = ArrayAdapter(this, android.R.layout.simple_spinner_item, tipos)
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item)
        spinnerTipo.adapter = adapter
    }

    private fun generarClaveDesdeUI() {
        val deviceCode = etDeviceCode.text.toString().trim().uppercase(Locale.ROOT)
        val tipo = spinnerTipo.selectedItem?.toString()?.uppercase(Locale.ROOT).orEmpty()

        if (deviceCode.isBlank()) {
            Toast.makeText(this, "El deviceCode es obligatorio", Toast.LENGTH_SHORT).show()
            etDeviceCode.requestFocus()
            return
        }

        if (tipo != "ANUAL" && tipo != "FULL") {
            Toast.makeText(this, "El tipo debe ser ANUAL o FULL", Toast.LENGTH_SHORT).show()
            return
        }

        etDeviceCode.setText(deviceCode)
        val clave = generarClave(deviceCode, tipo)
        tvResultado.text = clave
    }

    fun generarClave(deviceCode: String, tipo: String): String {
        val secreto = "PRST_APP_LICENSE_PRIVATE_2026_#84KF92@A"
        val input = deviceCode + tipo + secreto

        val hashBytes = MessageDigest.getInstance("SHA-256").digest(input.toByteArray())
        val hashHex = hashBytes.joinToString(separator = "") { "%02x".format(it) }.uppercase(Locale.ROOT)

        val primeros16 = hashHex.substring(0, 16)
        val bloques = primeros16.chunked(4).joinToString("-")

        return "$tipo-$bloques"
    }

    private fun copiarResultado() {
        val resultado = tvResultado.text.toString().trim()

        if (resultado.isBlank()) {
            Toast.makeText(this, "No hay clave para copiar", Toast.LENGTH_SHORT).show()
            return
        }

        val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
        clipboard.setPrimaryClip(ClipData.newPlainText("Clave de activación", resultado))
        Toast.makeText(this, "Clave copiada", Toast.LENGTH_SHORT).show()
    }

    private fun limpiarCampos() {
        etDeviceCode.text?.clear()
        spinnerTipo.setSelection(0)
        tvResultado.text = ""
        etDeviceCode.requestFocus()
    }
}
