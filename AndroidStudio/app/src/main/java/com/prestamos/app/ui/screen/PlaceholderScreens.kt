package com.prestamos.app.ui.screen

import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier

@Composable
fun ClientesScreen() = PlaceholderScreen(
    title = "Módulo de Clientes",
    description = "Fase 2: alta, edición, búsqueda y listado de clientes"
)

@Composable
fun PrestamosScreen() = PlaceholderScreen(
    title = "Módulo de Préstamos",
    description = "Fase 3: registro de préstamos y generación automática de cuotas"
)

@Composable
fun PagosScreen() = PlaceholderScreen(
    title = "Módulo de Pagos",
    description = "Fase 4: aplicación de abonos por cuota"
)

@Composable
fun ReportesScreen() = PlaceholderScreen(
    title = "Módulo de Reportes",
    description = "Fase 5: pendientes, vencidos y resumen general"
)

@Composable
private fun PlaceholderScreen(title: String, description: String) {
    Column(
        modifier = Modifier.fillMaxSize(),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement = Arrangement.Center
    ) {
        Text(text = title, style = MaterialTheme.typography.headlineSmall)
        Text(text = description, style = MaterialTheme.typography.bodyMedium)
    }
}
