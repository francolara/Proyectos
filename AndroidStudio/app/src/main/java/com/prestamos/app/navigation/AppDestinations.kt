package com.prestamos.app.navigation

enum class AppDestinations(val route: String, val title: String) {
    DASHBOARD("dashboard", "Inicio"),
    CLIENTES("clientes", "Clientes"),
    PRESTAMOS("prestamos", "Préstamos"),
    PAGOS("pagos", "Pagos"),
    REPORTES("reportes", "Reportes")
}
