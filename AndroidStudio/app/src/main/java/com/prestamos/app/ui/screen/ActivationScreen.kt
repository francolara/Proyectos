package com.prestamos.app.ui.screen

import android.content.ActivityNotFoundException
import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.widget.Toast
import androidx.compose.foundation.BorderStroke
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.KeyboardOptions
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.contentColorFor
import androidx.compose.material3.Icon
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedButton
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Text
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Email
import androidx.compose.material.icons.outlined.Message
import androidx.compose.runtime.Composable
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.luminance
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.KeyboardType
import androidx.compose.ui.unit.dp
import com.prestamos.app.data.license.LicenseType
import com.prestamos.app.ui.viewmodel.ActivationUiState
import com.prestamos.app.util.toDateString

private const val LICENSE_WHATSAPP_NUMBER = "51950305708"
private const val LICENSE_SUPPORT_EMAIL = "controlprestamos.app@gmail.com"

// Firma Codex 2026-03-21

@Composable
fun ActivationScreen(
    uiState: ActivationUiState,
    onActivationKeyChanged: (String) -> Unit,
    onActivate: () -> Unit
) {
    val context = LocalContext.current
    val status = uiState.status
    val dark = MaterialTheme.colorScheme.background.luminance() < 0.5f
    val requestPlanOptions = remember {
        listOf(
            RequestPlanCardUi(
                type = LicenseType.MENSUAL,
                title = "Mensual",
                primaryPrice = "S/ 10",
                usdReference = "≈ $ 2.70 USD",
                period = "30 dias",
                benefits = listOf("Acceso completo", "Todas las funciones")
            ),
            RequestPlanCardUi(
                type = LicenseType.ANUAL,
                title = "Anual",
                primaryPrice = "S/ 80",
                usdReference = "≈ $ 22 USD",
                period = "12 meses",
                benefits = listOf("Acceso completo", "Todas las funciones", "Mejor valor"),
                recommended = true
            ),
            RequestPlanCardUi(
                type = LicenseType.FULL,
                title = "Full",
                primaryPrice = "S/ 180",
                usdReference = "≈ $ 50 USD",
                period = "Pago unico",
                benefits = listOf("Acceso permanente", "Sin renovaciones")
            )
        )
    }
    var selectedRequestPlan by remember { mutableStateOf<LicenseType?>(null) }

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
                colors = CardDefaults.cardColors(
                    containerColor = if (dark) Color(0xFF1F2D21) else Color(0xFFEFF8EF)
                )
            ) {
                Column(modifier = Modifier.padding(14.dp), verticalArrangement = Arrangement.spacedBy(4.dp)) {
                    Text(
                        "Periodo de prueba activo",
                        style = MaterialTheme.typography.titleSmall,
                        color = MaterialTheme.colorScheme.onSurface
                    )
                    Text(
                        "Te quedan ${status.trialDaysRemaining} dia(s) para activar la version Pro.",
                        style = MaterialTheme.typography.bodyMedium,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }
        }

        Card(
            shape = RoundedCornerShape(18.dp),
            colors = CardDefaults.cardColors(
                containerColor = if (dark) Color(0xFF1F1F2A) else MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.45f)
            )
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
                    ),
                    color = MaterialTheme.colorScheme.onSurface
                )
                Row(modifier = Modifier.fillMaxWidth(), horizontalArrangement = Arrangement.End) {
                    OutlinedButton(onClick = {
                        if (copyToClipboard(context, status.deviceCode)) {
                            Toast.makeText(context, "Codigo copiado", Toast.LENGTH_SHORT).show()
                        }
                    }) {
                        Text("Copiar codigo")
                    }
                }
            }
        }

        Card(
            shape = RoundedCornerShape(18.dp),
            colors = CardDefaults.cardColors(
                containerColor = if (dark) Color(0xFF1F2D21) else Color(0xFFEFF8EF)
            )
        ) {
            val requestButtonTextStyle = MaterialTheme.typography.labelLarge.copy(
                fontSize = MaterialTheme.typography.labelLarge.fontSize * 0.85f
            )
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(12.dp)
            ) {
                Text(
                    text = "Solicitar licencia",
                    style = MaterialTheme.typography.titleMedium.copy(fontWeight = FontWeight.SemiBold)
                )
                Text(
                    text = "Copia tu codigo y envialo para activar tu licencia.",
                    style = MaterialTheme.typography.bodyMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
                Text("Plan", style = MaterialTheme.typography.labelLarge)
                Column(verticalArrangement = Arrangement.spacedBy(10.dp)) {
                    requestPlanOptions.forEach { option ->
                        val isSelected = selectedRequestPlan == option.type
                        val selectedContainer = if (dark) Color(0xFF2B4A2F) else Color(0xFFDDF2DE)
                        val normalContainer = if (dark) Color(0xFF1B2820) else Color(0xFFF3FAF3)
                        val borderColor = if (isSelected) MaterialTheme.colorScheme.primary else MaterialTheme.colorScheme.outline.copy(alpha = 0.45f)
                        Card(
                            shape = RoundedCornerShape(14.dp),
                            colors = CardDefaults.cardColors(containerColor = if (isSelected) selectedContainer else normalContainer),
                            border = BorderStroke(1.4.dp, borderColor),
                            modifier = Modifier
                                .fillMaxWidth()
                                .clickable { selectedRequestPlan = option.type }
                        ) {
                            Column(
                                modifier = Modifier.padding(14.dp),
                                verticalArrangement = Arrangement.spacedBy(5.dp)
                            ) {
                                Row(
                                    modifier = Modifier.fillMaxWidth(),
                                    horizontalArrangement = Arrangement.SpaceBetween
                                ) {
                                    Text(
                                        option.title,
                                        style = MaterialTheme.typography.titleMedium.copy(fontWeight = FontWeight.SemiBold)
                                    )
                                    if (option.recommended) {
                                        Text(
                                            "⭐ Recomendado",
                                            style = MaterialTheme.typography.labelMedium,
                                            color = MaterialTheme.colorScheme.primary
                                        )
                                    }
                                }
                                Text(
                                    option.primaryPrice,
                                    style = MaterialTheme.typography.headlineSmall.copy(fontWeight = FontWeight.Bold)
                                )
                                Text(
                                    option.usdReference,
                                    style = MaterialTheme.typography.bodySmall,
                                    color = MaterialTheme.colorScheme.onSurfaceVariant
                                )
                                Text(option.period, style = MaterialTheme.typography.labelMedium)
                                option.benefits.forEach { benefit ->
                                    Text("• $benefit", style = MaterialTheme.typography.bodySmall)
                                }
                            }
                        }
                    }
                }
                Text(
                    "*Precios referenciales en USD",
                    style = MaterialTheme.typography.bodySmall,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
                Row(
                    modifier = Modifier.fillMaxWidth(),
                    horizontalArrangement = Arrangement.spacedBy(10.dp)
                ) {
                    Button(
                        onClick = {
                            val selectedPlan = selectedRequestPlan
                            if (selectedPlan == null) {
                                Toast.makeText(context, "Selecciona un plan antes de continuar", Toast.LENGTH_SHORT).show()
                            } else {
                                openWhatsAppForLicense(context, status.deviceCode, selectedPlan.toRequestPlanLabel())
                            }
                        },
                        shape = RoundedCornerShape(12.dp),
                        modifier = Modifier
                            .weight(1f)
                            .height(44.dp)
                    ) {
                        Icon(
                            imageVector = Icons.Outlined.Message,
                            contentDescription = null
                        )
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("WhatsApp", style = requestButtonTextStyle)
                    }
                    OutlinedButton(
                        onClick = {
                            val selectedPlan = selectedRequestPlan
                            if (selectedPlan == null) {
                                Toast.makeText(context, "Selecciona un plan antes de continuar", Toast.LENGTH_SHORT).show()
                            } else {
                                openEmailForLicense(context, status.deviceCode, selectedPlan.toRequestPlanLabel())
                            }
                        },
                        shape = RoundedCornerShape(12.dp),
                        modifier = Modifier
                            .weight(1f)
                            .height(44.dp)
                    ) {
                        Icon(
                            imageVector = Icons.Outlined.Email,
                            contentDescription = null
                        )
                        Spacer(modifier = Modifier.width(8.dp))
                        Text("Correo", style = requestButtonTextStyle)
                    }
                }
            }
        }

        val showActivationForm = status.licenseType == LicenseType.TRIAL || !status.isActivated || !status.isValid
        if (showActivationForm) {
            Card(
                shape = RoundedCornerShape(18.dp),
                colors = CardDefaults.cardColors(
                    containerColor = if (dark) Color(0xFF1C2226) else MaterialTheme.colorScheme.surface
                )
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
            status.manipulatedDateDetected || (status.licenseType == LicenseType.TRIAL && status.trialExpired) -> if (dark) Color(0xFF3A2323) else Color(0xFFFFECEC)
            status.isValid && status.isActivated -> if (dark) Color(0xFF1F3021) else Color(0xFFEAF7EA)
            else -> if (dark) Color(0xFF352D1F) else Color(0xFFFFF7E6)
        }
        val statusTextColor = contentColorFor(statusContainerColor)

        Card(
            shape = RoundedCornerShape(18.dp),
            colors = CardDefaults.cardColors(
                containerColor = statusContainerColor,
                contentColor = statusTextColor
            )
        ) {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(16.dp),
                verticalArrangement = Arrangement.spacedBy(6.dp)
            ) {
                Text(
                    statusTitle,
                    style = MaterialTheme.typography.titleLarge.copy(fontWeight = FontWeight.Bold)
                )
                Text(
                    text = "Plan",
                    style = MaterialTheme.typography.labelMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
                Text(planText, style = MaterialTheme.typography.bodyLarge.copy(fontWeight = FontWeight.Medium))
                Text(
                    text = "Valida hasta",
                    style = MaterialTheme.typography.labelMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
                Text(vigenciaText, style = MaterialTheme.typography.bodyLarge.copy(fontWeight = FontWeight.Medium))
                if (status.manipulatedDateDetected) {
                    Text(
                        "Se detecto un cambio de fecha en el dispositivo. Verifica la fecha y reactiva la licencia.",
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.error
                    )
                }
            }
        }

    }
}

private fun copyToClipboard(context: Context, value: String): Boolean {
    val clipboard = context.getSystemService(Context.CLIPBOARD_SERVICE) as? ClipboardManager ?: return false
    clipboard.setPrimaryClip(ClipData.newPlainText("device_code", value))
    return true
}

private fun openWhatsAppForLicense(context: Context, deviceCode: String, selectedPlan: String) {
    val message = "Hola, quiero activar la app.\nPlan: $selectedPlan\nCodigo de dispositivo: $deviceCode"
    val uri = Uri.parse("https://wa.me/$LICENSE_WHATSAPP_NUMBER?text=${Uri.encode(message)}")
    val intent = Intent(Intent.ACTION_VIEW, uri)
    val packageManager = context.packageManager
    val hasWhatsApp = packageManager.getLaunchIntentForPackage("com.whatsapp") != null ||
        packageManager.getLaunchIntentForPackage("com.whatsapp.w4b") != null

    if (!hasWhatsApp) {
        Toast.makeText(context, "WhatsApp no esta instalado", Toast.LENGTH_SHORT).show()
        return
    }

    try {
        context.startActivity(intent)
    } catch (_: ActivityNotFoundException) {
        Toast.makeText(context, "WhatsApp no esta instalado", Toast.LENGTH_SHORT).show()
    }
}

private fun openEmailForLicense(context: Context, deviceCode: String, selectedPlan: String) {
    val subject = "Solicitud de licencia"
    val body = "Hola, quiero solicitar una licencia.\nPlan: $selectedPlan\nCodigo de dispositivo: $deviceCode"
    val intent = Intent(Intent.ACTION_SENDTO).apply {
        data = Uri.parse("mailto:")
        putExtra(Intent.EXTRA_EMAIL, arrayOf(LICENSE_SUPPORT_EMAIL))
        putExtra(Intent.EXTRA_SUBJECT, subject)
        putExtra(Intent.EXTRA_TEXT, body)
    }

    try {
        context.startActivity(intent)
    } catch (_: ActivityNotFoundException) {
        Toast.makeText(context, "No se encontro una app de correo", Toast.LENGTH_SHORT).show()
    }
}

private fun LicenseType.toRequestPlanLabel(): String = when (this) {
    LicenseType.MENSUAL -> "Mensual"
    LicenseType.ANUAL -> "Anual"
    LicenseType.FULL -> "Full"
    LicenseType.TRIAL -> "Prueba"
}

private data class RequestPlanCardUi(
    val type: LicenseType,
    val title: String,
    val primaryPrice: String,
    val usdReference: String,
    val period: String,
    val benefits: List<String>,
    val recommended: Boolean = false
)
