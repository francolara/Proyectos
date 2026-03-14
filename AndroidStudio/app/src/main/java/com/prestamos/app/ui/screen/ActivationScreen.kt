package com.prestamos.app.ui.screen

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.material3.Button
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import com.prestamos.app.data.license.LicenseType
import com.prestamos.app.ui.viewmodel.ActivationUiState
import com.prestamos.app.util.toDateString

@Composable
fun ActivationScreen(
    uiState: ActivationUiState,
    onActivationKeyChanged: (String) -> Unit,
    onActivate: () -> Unit,
    onRefresh: () -> Unit
) {
    val context = LocalContext.current
    val status = uiState.status

    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(12.dp)
    ) {
        Text("Activación de licencia", style = MaterialTheme.typography.headlineSmall)
        Text("La app funciona 30 días en modo prueba. Luego requiere activación manual.")

        Text("Código del equipo: ${status.deviceCode}", style = MaterialTheme.typography.titleMedium)
        Button(onClick = { copyToClipboard(context, status.deviceCode) }) {
            Text("Copiar código")
        }

        OutlinedTextField(
            value = uiState.activationKey,
            onValueChange = onActivationKeyChanged,
            label = { Text("Clave de activación") },
            singleLine = true,
            keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Ascii),
            modifier = Modifier.fillMaxWidth()
        )

        Button(onClick = onActivate, modifier = Modifier.fillMaxWidth()) {
            Text("Activar")
        }

        when {
            status.manipulatedDateDetected -> {
                Text("Se detectó manipulación de fecha del sistema. Reactiva la licencia.")
            }

            status.licenseType == LicenseType.TRIAL && !status.trialExpired -> {
                Text("Trial activo. Días restantes: ${status.trialDaysRemaining}", color = Color.Red)
            }

            status.licenseType == LicenseType.TRIAL && status.trialExpired -> {
                Text("El trial expiró. Debes activar la app para continuar.")
            }

            status.licenseType == LicenseType.ANUAL && status.expirationDate != null -> {
                Text("Licencia ANUAL activa hasta: ${status.expirationDate.toDateString()}")
            }

            status.licenseType == LicenseType.FULL -> {
                Text("Licencia FULL activa (sin vencimiento)")
            }
        }

        Button(onClick = onRefresh) {
            Text("Revalidar estado")
        }
    }
}

private fun copyToClipboard(context: Context, value: String) {
    val clipboard = context.getSystemService(Context.CLIPBOARD_SERVICE) as? ClipboardManager ?: return
    clipboard.setPrimaryClip(ClipData.newPlainText("device_code", value))
}
