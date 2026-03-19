package com.prestamos.app.ui.screen

import androidx.compose.foundation.background
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
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
import androidx.compose.foundation.verticalScroll
import androidx.compose.material3.Button
import androidx.compose.material3.Card
import androidx.compose.material3.CardDefaults
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.OutlinedTextField
import androidx.compose.material3.Surface
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.getValue
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.window.Dialog
import com.prestamos.app.ui.viewmodel.OnboardingUiState

private data class CurrencyOption(
    val code: String,
    val symbol: String,
    val name: String
) {
    val displayText: String get() = "$symbol - $name"
}

private val currencyOptions = listOf(
    CurrencyOption("PEN", "S/", "Sol peruano"),
    CurrencyOption("USD", "$", "Dolar estadounidense"),
    CurrencyOption("ARS", "$", "Peso argentino"),
    CurrencyOption("BOB", "Bs", "Boliviano"),
    CurrencyOption("BRL", "R$", "Real brasileno"),
    CurrencyOption("CLP", "$", "Peso chileno"),
    CurrencyOption("COP", "$", "Peso colombiano"),
    CurrencyOption("PYG", "Gs", "Guarani"),
    CurrencyOption("UYU", "\$U", "Peso uruguayo"),
    CurrencyOption("VES", "Bs", "Bolivar venezolano"),
    CurrencyOption("GYD", "G$", "Dolar guyanes"),
    CurrencyOption("SRD", "$", "Dolar surinames")
)

private val primaryCurrencyCodes = setOf("PEN", "USD")

@Composable
fun OnboardingScreen(
    uiState: OnboardingUiState,
    onComenzar: () -> Unit,
    onBusinessNameChange: (String) -> Unit,
    onMainCurrencySelected: (String) -> Unit,
    onSecondaryCurrencySelected: (String?) -> Unit,
    onFinalizar: () -> Unit
) {
    var showMainCurrencyPicker by remember { mutableStateOf(false) }
    var showSecondaryCurrencyPicker by remember { mutableStateOf(false) }

    Box(
        modifier = Modifier
            .fillMaxSize()
            .background(MaterialTheme.colorScheme.background)
            .padding(20.dp),
        contentAlignment = Alignment.Center
    ) {
        if (uiState.step == 0) {
            Card(
                shape = RoundedCornerShape(24.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                elevation = CardDefaults.cardElevation(defaultElevation = 4.dp),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 20.dp, vertical = 24.dp),
                    horizontalAlignment = Alignment.CenterHorizontally,
                    verticalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    Text(
                        text = "Bienvenido",
                        style = MaterialTheme.typography.headlineSmall.copy(fontWeight = FontWeight.SemiBold)
                    )
                    Text(
                        text = "Configura tu app en menos de 1 minuto",
                        style = MaterialTheme.typography.bodyMedium,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                    Spacer(modifier = Modifier.height(6.dp))
                    Button(
                        onClick = onComenzar,
                        shape = RoundedCornerShape(14.dp),
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Text("Comenzar")
                    }
                }
            }
        } else {
            Card(
                shape = RoundedCornerShape(24.dp),
                colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface),
                elevation = CardDefaults.cardElevation(defaultElevation = 4.dp),
                modifier = Modifier.fillMaxWidth()
            ) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .verticalScroll(rememberScrollState())
                        .padding(horizontal = 18.dp, vertical = 18.dp),
                    verticalArrangement = Arrangement.spacedBy(12.dp)
                ) {
                    Text(
                        text = "Datos basicos",
                        style = MaterialTheme.typography.titleLarge.copy(fontWeight = FontWeight.SemiBold)
                    )

                    OutlinedTextField(
                        value = uiState.businessName,
                        onValueChange = onBusinessNameChange,
                        label = { Text("Nombre del negocio o prestamista") },
                        singleLine = true,
                        shape = RoundedCornerShape(14.dp),
                        modifier = Modifier.fillMaxWidth()
                    )

                    CurrencyField(
                        label = "Moneda principal",
                        value = uiState.mainCurrencyCode.toCurrencyDisplay(),
                        placeholder = "Selecciona una moneda",
                        onClick = { showMainCurrencyPicker = true }
                    )

                    CurrencyField(
                        label = "Moneda secundaria (opcional)",
                        value = uiState.secondaryCurrencyCode.toCurrencyDisplay(),
                        placeholder = "Opcional",
                        onClick = { showSecondaryCurrencyPicker = true }
                    )

                    if (!uiState.errorMessage.isNullOrBlank()) {
                        Text(
                            text = uiState.errorMessage,
                            color = MaterialTheme.colorScheme.error,
                            style = MaterialTheme.typography.bodySmall
                        )
                    }

                    Button(
                        onClick = onFinalizar,
                        shape = RoundedCornerShape(14.dp),
                        modifier = Modifier.fillMaxWidth()
                    ) {
                        Text("Continuar")
                    }
                }
            }
        }
    }

    if (showMainCurrencyPicker) {
        CurrencyPickerDialog(
            title = "Seleccionar moneda principal",
            selectedCode = uiState.mainCurrencyCode,
            allowNone = false,
            onDismiss = { showMainCurrencyPicker = false },
            onSelect = {
                onMainCurrencySelected(it ?: return@CurrencyPickerDialog)
                showMainCurrencyPicker = false
            }
        )
    }

    if (showSecondaryCurrencyPicker) {
        CurrencyPickerDialog(
            title = "Seleccionar moneda secundaria",
            selectedCode = uiState.secondaryCurrencyCode,
            allowNone = true,
            onDismiss = { showSecondaryCurrencyPicker = false },
            onSelect = {
                onSecondaryCurrencySelected(it)
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
                .clickable { onClick() },
            shape = RoundedCornerShape(14.dp),
            border = androidx.compose.foundation.BorderStroke(1.dp, MaterialTheme.colorScheme.outlineVariant),
            color = Color.Transparent
        ) {
            Text(
                text = if (value.isBlank()) placeholder else value,
                modifier = Modifier.padding(horizontal = 14.dp, vertical = 14.dp),
                style = MaterialTheme.typography.bodyLarge,
                color = if (value.isBlank()) MaterialTheme.colorScheme.onSurfaceVariant else MaterialTheme.colorScheme.onSurface
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
    val primary = currencyOptions.filter { it.code in primaryCurrencyCodes }
    val others = currencyOptions.filterNot { it.code in primaryCurrencyCodes }

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

private fun String?.toCurrencyDisplay(): String {
    if (this.isNullOrBlank()) return ""
    return currencyOptions.firstOrNull { it.code == this }?.displayText.orEmpty()
}
