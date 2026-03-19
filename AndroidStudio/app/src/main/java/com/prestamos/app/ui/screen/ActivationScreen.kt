package com.prestamos.app.ui.screen

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
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
            .verticalScroll(rememberScrollState())
            .padding(horizontal = 18.dp, vertical = 16.dp),
        verticalArrangement = Arrangement.spacedBy(14.dp)
    ) {
        Text(
            text = "Activar version Pro",
            style = MaterialTheme.typography.headlineSmall.copy(fontWeight = FontWeight.SemiBold)
        )
        Text(
            text = "Activa tu version Pro para seguir usando todas las funciones de la app.",
            style = MaterialTheme.typography.bodyMedium,
            color = MaterialTheme.colorScheme.onSurfaceVariant
        )

        if (status.licenseType == LicenseType.TRIAL && !status.trialExpired) {
            Card(
                shape = RoundedCornerShape(16.dp),
                colors = CardDefaults.cardColors(containerColor = Color(0xFFEFF8EF))
            ) {
                Column(modifier = Modifier.padding(14.dp), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                    Text("Periodo de prueba activo", style = MaterialTheme.typography.titleSmall)
                    Text(
                        "Te quedan ${status.trialDaysRemaining} dia(s) para activar la version Pro.",
                        style = MaterialTheme.typography.bodyMedium
                    )
                }
            }
        }

        Card(
            shape = RoundedCornerShape(18.dp),
            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.45f))
        ) {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(10.dp)
            ) {
                Text("Tu codigo de dispositivo", style = MaterialTheme.typography.labelLarge)
                Text(
                    text = status.deviceCode,
                    style = MaterialTheme.typography.titleLarge.copy(
                        fontFamily = FontFamily.Monospace,
                        fontWeight = FontWeight.Bold
                    )
                )
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.End) {
                    OutlinedButton(onClick = { copyToClipboard(context, status.deviceCode) }) {
                        Text("Copiar codigo")
                    }
                }
            }
        }

        val showActivationForm = status.licenseType == LicenseType.TRIAL || !status.isActivated || !status.isValid
        if (showActivationForm) {
            Card(
                shape = RoundedCornerShape(18.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface)
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(16.dp),
                    verticalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    OutlinedTextField(
                        value = uiState.activationKey,
                        onValueChange = onActivationKeyChanged,
                        label = { Text("Codigo de activacion") },
                        placeholder = { Text("Ingresa o pega tu codigo") },
                        singleLine = true,
                        keyboardOptions = KeyboardOptions(keyboardType = KeyboardType.Ascii),
                        shape = RoundedCornerShape(14.dp),
                        modifier = Modifier.fillMaxWidth()
                    )

                    Button(
                        onClick = onActivate,
                        enabled = !uiState.loading,
                        shape = RoundedCornerShape(14.dp),
                        modifier = Modifier
                            .fillMaxWidth()
                    ) {
                        if (uiState.loading) {
                            CircularProgressIndicator(
                                modifier = Modifier.padding(vertical = 2.dp),
                                strokeWidth = 2.dp
                            )
                        } else {
                            Text("Activar licencia", modifier = Modifier.padding(vertical = 4.dp))
                        }
                    }
                }
            }
        }

        val statusTitle = when {
            status.manipulatedDateDetected -> "Estado de licencia"
            status.licenseType == LicenseType.TRIAL && status.trialExpired -> "Licencia inactiva"
            status.licenseType == LicenseType.TRIAL -> "Periodo de prueba"
            status.isValid && status.isActivated -> "Licencia activa"
            else -> "Licencia inactiva"
        }

        val planText = when (status.licenseType) {
            LicenseType.MENSUAL -> "Mensual"
            LicenseType.ANUAL -> "Anual"
            LicenseType.FULL -> "Full"
            LicenseType.TRIAL -> "Prueba"
        }

        val vigenciaText = when {
            status.licenseType == LicenseType.FULL && status.isActivated -> "Sin vencimiento"
            status.expirationDate != null -> status.expirationDate.toDateString()
            status.licenseType == LicenseType.TRIAL && !status.trialExpired -> "${status.trialDaysRemaining} dia(s) restantes"
            else -> "No disponible"
        }

        val statusContainerColor = when {
            status.manipulatedDateDetected || (status.licenseType == LicenseType.TRIAL && status.trialExpired) -> Color(0xFFFFECEC)
            status.isValid && status.isActivated -> Color(0xFFEAF7EA)
            else -> Color(0xFFFFF7E6)
        }

        Card(
            shape = RoundedCornerShape(18.dp),
            colors = CardDefaults.cardColors(containerColor = statusContainerColor)
        ) {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(6.dp)
            ) {
                Text(statusTitle, style = MaterialTheme.typography.titleMedium.copy(fontWeight = FontWeight.SemiBold))
                Text("Plan: $planText", style = MaterialTheme.typography.bodyMedium)
                Text("Valida hasta: $vigenciaText", style = MaterialTheme.typography.bodyMedium)
                if (status.manipulatedDateDetected) {
                    Text(
                        "Se detecto un cambio de fecha en el dispositivo. Verifica la fecha y reactiva la licencia.",
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.error
                    )
                }
            }
        }

        OutlinedButton(
            onClick = onRefresh,
            shape = RoundedCornerShape(12.dp),
            modifier = Modifier.align(Alignment.End)
        ) {
            Text("Verificar licencia")
        }
    }
}

private fun copyToClipboard(context: Context, value: String) {
    val clipboard = context.getSystemService(Context.CLIPBOARD_SERVICE) as? ClipboardManager ?: return
    clipboard.setPrimaryClip(ClipData.newPlainText("device_code", value))
}
