package com.prestamos.app

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.os.Bundle
import android.text.Editable
import android.text.InputFilter
import android.text.TextWatcher
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

    companion object {
        private const val LICENSE_SECRET = "PRST_APP_LICENSE_PRIVATE_2026_#84KF92@A"
        private val TIPOS_VALIDOS = setOf("ANUAL", "FULL")
    }

    private lateinit var etDeviceCode: EditText
    private lateinit var spinnerTipo: Spinner
    private lateinit var tvResultadoValor: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        etDeviceCode = findViewById(R.id.etDeviceCode)
        spinnerTipo = findViewById(R.id.spinnerTipo)
        tvResultadoValor = findViewById(R.id.tvResultadoValor)
        val btnGenerar: Button = findViewById(R.id.btnGenerar)
        val btnCopiar: Button = findViewById(R.id.btnCopiar)
        val btnLimpiar: Button = findViewById(R.id.btnLimpiar)

        configurarTipoLicencia()
        configurarMayusculasAutomaticas()

        btnGenerar.setOnClickListener { generarYMostrarClave() }
        btnCopiar.setOnClickListener { copiarClave() }
        btnLimpiar.setOnClickListener { limpiarCampos() }
    }

    private fun configurarTipoLicencia() {
        val tipos = listOf("ANUAL", "FULL")
        val adapter = ArrayAdapter(this, android.R.layout.simple_spinner_item, tipos)
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item)
        spinnerTipo.adapter = adapter
        spinnerTipo.setSelection(0)
    }

    private fun configurarMayusculasAutomaticas() {
        etDeviceCode.filters = arrayOf(InputFilter.AllCaps())
        etDeviceCode.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) = Unit

            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) = Unit

            override fun afterTextChanged(s: Editable?) {
                val textoActual = s?.toString().orEmpty()
                val textoMayuscula = textoActual.uppercase(Locale.ROOT)
                if (textoActual != textoMayuscula) {
                    etDeviceCode.setText(textoMayuscula)
                    etDeviceCode.setSelection(textoMayuscula.length)
                }
            }
        })
    }

    private fun generarYMostrarClave() {
        val deviceCode = etDeviceCode.text.toString().trim().uppercase(Locale.ROOT)
        val tipo = spinnerTipo.selectedItem.toString()

        if (deviceCode.isEmpty()) {
            etDeviceCode.error = getString(R.string.error_device_code_vacio)
            etDeviceCode.requestFocus()
            return
        }

        val clave = generarClave(deviceCode, tipo)
        tvResultadoValor.text = clave
    }

    private fun copiarClave() {
        val clave = tvResultadoValor.text.toString().trim()
        if (clave.isEmpty()) {
            Toast.makeText(this, getString(R.string.error_sin_clave), Toast.LENGTH_SHORT).show()
            return
        }

        val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
        val clip = ClipData.newPlainText(getString(R.string.etiqueta_clave_generada), clave)
        clipboard.setPrimaryClip(clip)
        Toast.makeText(this, getString(R.string.clave_copiada), Toast.LENGTH_SHORT).show()
    }

    private fun limpiarCampos() {
        etDeviceCode.text.clear()
        spinnerTipo.setSelection(0)
        tvResultadoValor.text = ""
        etDeviceCode.error = null
    }

    fun generarClave(deviceCode: String, tipo: String): String {
        val tipoNormalizado = tipo.uppercase(Locale.ROOT)
        require(tipoNormalizado in TIPOS_VALIDOS) { "Tipo de licencia inválido. Solo se permite ANUAL o FULL." }

        val texto = deviceCode.uppercase(Locale.ROOT) + tipoNormalizado + LICENSE_SECRET
        val hash = sha256(texto)
        val primeros16 = hash.take(16).uppercase(Locale.ROOT)
        val bloques = primeros16.chunked(4).joinToString("-")
        return "$tipoNormalizado-$bloques"
    }

    private fun sha256(texto: String): String {
        val digest = MessageDigest.getInstance("SHA-256")
        val hash = digest.digest(texto.toByteArray(Charsets.UTF_8))
        return hash.joinToString(separator = "") { "%02x".format(it) }
    }
}
