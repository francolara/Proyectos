package com.prestamos.app.ui.screen

import android.widget.Toast
import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.Dialog
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import com.prestamos.app.data.config.InitialSetupPreferences
import com.prestamos.app.ui.viewmodel.AppViewModel

private data class ConfigCurrencyOption(
    val code: String,
    val symbol: String,
    val name: String
) {
    val displayText: String get() = "$symbol - $name"
}

private val configCurrencyOptions = listOf(
    ConfigCurrencyOption("PEN", "S/", "Sol peruano"),
    ConfigCurrencyOption("USD", "$", "Dolar estadounidense"),
    ConfigCurrencyOption("ARS", "$", "Peso argentino"),
    ConfigCurrencyOption("BOB", "Bs", "Boliviano"),
    ConfigCurrencyOption("BRL", "R$", "Real brasileno"),
    ConfigCurrencyOption("CLP", "$", "Peso chileno"),
    ConfigCurrencyOption("COP", "$", "Peso colombiano"),
    ConfigCurrencyOption("PYG", "Gs", "Guarani"),
    ConfigCurrencyOption("UYU", "\$U", "Peso uruguayo"),
    ConfigCurrencyOption("VES", "Bs", "Bolivar venezolano"),
    ConfigCurrencyOption("GYD", "G$", "Dolar guyanes"),
    ConfigCurrencyOption("SRD", "$", "Dolar surinames")
)

private val primaryCurrencyCodes = setOf("PEN", "USD")

@Composable
fun ConfiguracionScreen(viewModel: AppViewModel) {
    val context = androidx.compose.ui.platform.LocalContext.current
    val prefs = remember { InitialSetupPreferences(context) }
    val prestamos by viewModel.prestamos.collectAsStateWithLifecycle()
    val usedCurrencyCodes = remember(prestamos) { prestamos.map { it.moneda.code.uppercase() }.toSet() }

    var businessName by remember { mutableStateOf("") }
    var mainCurrencyCode by remember { mutableStateOf<String?>(null) }
    var secondaryCurrencyCode by remember { mutableStateOf<String?>(null) }
    var originalMainCurrencyCode by remember { mutableStateOf<String?>(null) }
    var originalSecondaryCurrencyCode by remember { mutableStateOf<String?>(null) }
    var showMainCurrencyPicker by remember { mutableStateOf(false) }
    var showSecondaryCurrencyPicker by remember { mutableStateOf(false) }
    var loaded by remember { mutableStateOf(false) }

    LaunchedEffect(Unit) {
        businessName = prefs.getBusinessName()
        mainCurrencyCode = prefs.getMainCurrencyCode()
        secondaryCurrencyCode = prefs.getSecondaryCurrencyCode()
        originalMainCurrencyCode = mainCurrencyCode
        originalSecondaryCurrencyCode = secondaryCurrencyCode
        loaded = true
    }

    if (!loaded) return

    Column(
        modifier = Modifier
            .verticalScroll(rememberScrollState())
            .padding(16.dp),
        verticalArrangement = Arrangement.spacedBy(12.dp)
    ) {
        Text("Configuracion", style = MaterialTheme.typography.headlineSmall)
        Card(
            shape = RoundedCornerShape(16.dp),
            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.45f)),
            modifier = Modifier.fillMaxWidth()
        ) {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(horizontal = 14.dp, vertical = 12.dp),
                verticalArrangement = Arrangement.spacedBy(10.dp)
            ) {
                androidx.compose.material3.OutlinedTextField(
                    value = businessName,
                    onValueChange = { businessName = it },
                    label = { Text("Nombre del Negocio") },
                    singleLine = true,
                    shape = RoundedCornerShape(12.dp),
                    modifier = Modifier.fillMaxWidth()
                )
                CurrencyField(
                    label = "Moneda principal",
                    value = mainCurrencyCode.toConfigCurrencyDisplay(),
                    placeholder = "Selecciona una moneda",
                    enabled = originalMainCurrencyCode.isNullOrBlank() || !usedCurrencyCodes.contains(originalMainCurrencyCode?.uppercase()),
                    helper = if (!originalMainCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalMainCurrencyCode?.uppercase())) {
                        "No se puede cambiar porque ya existen prestamos en esa moneda"
                    } else {
                        null
                    },
                    onClick = {
                        val mainLocked = !originalMainCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalMainCurrencyCode?.uppercase())
                        if (!mainLocked) showMainCurrencyPicker = true
                    }
                )
                CurrencyField(
                    label = "Moneda secundaria (opcional)",
                    value = secondaryCurrencyCode.toConfigCurrencyDisplay(),
                    placeholder = "Opcional",
                    enabled = originalSecondaryCurrencyCode.isNullOrBlank() || !usedCurrencyCodes.contains(originalSecondaryCurrencyCode?.uppercase()),
                    helper = if (!originalSecondaryCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalSecondaryCurrencyCode?.uppercase())) {
                        "No se puede cambiar porque ya existen prestamos en esa moneda"
                    } else {
                        null
                    },
                    onClick = {
                        val secondaryLocked = !originalSecondaryCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalSecondaryCurrencyCode?.uppercase())
                        if (!secondaryLocked) showSecondaryCurrencyPicker = true
                    }
                )
                Button(
                    onClick = {
                        val name = businessName.trim()
                        val main = mainCurrencyCode
                        if (name.isBlank()) {
                            Toast.makeText(context, "Ingresa el nombre del negocio o prestamista", Toast.LENGTH_SHORT).show()
                            return@Button
                        }
                        if (main.isNullOrBlank()) {
                            Toast.makeText(context, "Selecciona una moneda principal", Toast.LENGTH_SHORT).show()
                            return@Button
                        }
                        val mainLocked = !originalMainCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalMainCurrencyCode?.uppercase())
                        if (mainLocked && main != originalMainCurrencyCode) {
                            Toast.makeText(context, "No se puede cambiar la moneda principal porque ya tiene prestamos creados", Toast.LENGTH_SHORT).show()
                            return@Button
                        }
                        val secondaryLocked = !originalSecondaryCurrencyCode.isNullOrBlank() && usedCurrencyCodes.contains(originalSecondaryCurrencyCode?.uppercase())
                        if (secondaryLocked && secondaryCurrencyCode != originalSecondaryCurrencyCode) {
                            Toast.makeText(context, "No se puede modificar la moneda secundaria porque ya tiene prestamos creados", Toast.LENGTH_SHORT).show()
                            return@Button
                        }
                        prefs.updateConfiguration(
                            businessName = name,
                            mainCurrencyCode = main,
                            secondaryCurrencyCode = secondaryCurrencyCode
                        )
                        Toast.makeText(context, "Configuracion actualizada", Toast.LENGTH_SHORT).show()
                    },
                    shape = RoundedCornerShape(12.dp),
                    modifier = Modifier.fillMaxWidth()
                ) {
                    Text("Guardar cambios")
                }
            }
        }
    }

    if (showMainCurrencyPicker) {
        CurrencyPickerDialog(
            title = "Seleccionar moneda principal",
            selectedCode = mainCurrencyCode,
            allowNone = false,
            onDismiss = { showMainCurrencyPicker = false },
            onSelect = {
                mainCurrencyCode = it ?: return@CurrencyPickerDialog
                if (secondaryCurrencyCode == mainCurrencyCode) secondaryCurrencyCode = null
                showMainCurrencyPicker = false
            }
        )
    }

    if (showSecondaryCurrencyPicker) {
        CurrencyPickerDialog(
            title = "Seleccionar moneda secundaria",
            selectedCode = secondaryCurrencyCode,
            allowNone = true,
            onDismiss = { showSecondaryCurrencyPicker = false },
            onSelect = {
                secondaryCurrencyCode = if (it == mainCurrencyCode) null else it
                showSecondaryCurrencyPicker = false
            }
        )
    }
}

@Composable
private fun CurrencyField(
    label: String,
    value: String,
    placeholder: String,
    enabled: Boolean,
    helper: String? = null,
    onClick: () -> Unit
) {
    Column(verticalArrangement = Arrangement.spacedBy(4.dp)) {
        Text(
            text = label,
            style = MaterialTheme.typography.labelLarge,
            color = MaterialTheme.colorScheme.onSurfaceVariant
        )
        Surface(
            modifier = Modifier
                .fillMaxWidth()
                .clickable(enabled = enabled) { onClick() },
            shape = RoundedCornerShape(14.dp),
            border = androidx.compose.foundation.BorderStroke(1.dp, MaterialTheme.colorScheme.outlineVariant),
            color = if (enabled) Color.Transparent else MaterialTheme.colorScheme.surfaceVariant.copy(alpha = 0.35f)
        ) {
            Text(
                text = if (value.isBlank()) placeholder else value,
                modifier = Modifier.padding(horizontal = 14.dp, vertical = 14.dp),
                style = MaterialTheme.typography.bodyLarge,
                color = if (value.isBlank()) MaterialTheme.colorScheme.onSurfaceVariant else MaterialTheme.colorScheme.onSurface
            )
        }
        if (!helper.isNullOrBlank()) {
            Text(
                text = helper,
                style = MaterialTheme.typography.bodySmall,
                color = MaterialTheme.colorScheme.onSurfaceVariant
            )
        }
    }
}

@Composable
private fun CurrencyPickerDialog(
    title: String,
    selectedCode: String?,
    allowNone: Boolean,
    onDismiss: () -> Unit,
    onSelect: (String?) -> Unit
) {
    val primary = configCurrencyOptions.filter { it.code in primaryCurrencyCodes }
    val others = configCurrencyOptions.filterNot { it.code in primaryCurrencyCodes }

    Dialog(onDismissRequest = onDismiss) {
        Card(
            shape = RoundedCornerShape(18.dp),
            modifier = Modifier.fillMaxWidth(),
            colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface)
        ) {
            Column(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(14.dp)
            ) {
                Text(title, style = MaterialTheme.typography.titleMedium.copy(fontWeight = FontWeight.SemiBold))
                Spacer(modifier = Modifier.height(10.dp))

                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .height(360.dp)
                        .verticalScroll(rememberScrollState()),
                    verticalArrangement = Arrangement.spacedBy(6.dp)
                ) {
                    if (allowNone) {
                        CurrencyOptionRow(
                            label = "Sin moneda secundaria",
                            selected = selectedCode == null,
                            onClick = { onSelect(null) }
                        )
                    }

                    Text("Monedas principales", style = MaterialTheme.typography.labelLarge)
                    primary.forEach { option ->
                        CurrencyOptionRow(
                            label = option.displayText,
                            code = option.code,
                            selected = selectedCode == option.code,
                            onClick = { onSelect(option.code) }
                        )
                    }

                    Spacer(modifier = Modifier.height(4.dp))
                    Text("Otras monedas", style = MaterialTheme.typography.labelLarge)
                    others.forEach { option ->
                        CurrencyOptionRow(
                            label = option.displayText,
                            code = option.code,
                            selected = selectedCode == option.code,
                            onClick = { onSelect(option.code) }
                        )
                    }
                }
            }
        }
    }
}

@Composable
private fun CurrencyOptionRow(
    label: String,
    code: String? = null,
    selected: Boolean,
    onClick: () -> Unit
) {
    Surface(
        modifier = Modifier
            .fillMaxWidth()
            .clickable { onClick() },
        shape = RoundedCornerShape(12.dp),
        color = if (selected) Color(0xFFE8F6E9) else Color.Transparent
    ) {
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 12.dp, vertical = 10.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            Box(
                modifier = Modifier
                    .width(8.dp)
                    .height(8.dp)
                    .background(
                        color = if (selected) Color(0xFF2E7D32) else Color(0xFFB0BEC5),
                        shape = RoundedCornerShape(50)
                    )
            )
            Spacer(modifier = Modifier.width(10.dp))
            Text(
                text = label,
                maxLines = 1,
                overflow = TextOverflow.Ellipsis,
                style = MaterialTheme.typography.bodyMedium,
                modifier = Modifier.weight(1f)
            )
            if (code != null) {
                Spacer(modifier = Modifier.width(8.dp))
                Text(
                    text = code,
                    style = MaterialTheme.typography.labelMedium,
                    color = MaterialTheme.colorScheme.onSurfaceVariant
                )
            }
        }
    }
}

private fun String?.toConfigCurrencyDisplay(): String {
    if (this.isNullOrBlank()) return ""
    return configCurrencyOptions.firstOrNull { it.code == this }?.displayText.orEmpty()
}
