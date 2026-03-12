package com.prestamos.app

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Assessment
import androidx.compose.material.icons.outlined.Groups
import androidx.compose.material.icons.outlined.Payments
import androidx.compose.material.icons.outlined.RequestPage
import androidx.compose.material3.Icon
import androidx.compose.material3.NavigationBar
import androidx.compose.material3.NavigationBarItem
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.ui.Modifier
import androidx.navigation.NavGraph.Companion.findStartDestination
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.currentBackStackEntryAsState
import androidx.navigation.compose.rememberNavController
import com.prestamos.app.data.local.AppDatabase
import com.prestamos.app.navigation.AppDestinations
import com.prestamos.app.ui.screen.ClientesScreen
import com.prestamos.app.ui.screen.PagosScreen
import com.prestamos.app.ui.screen.PrestamosScreen
import com.prestamos.app.ui.screen.ReportesScreen
import com.prestamos.app.ui.theme.AppPrestamosTheme

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        AppDatabase.getInstance(applicationContext)
        enableEdgeToEdge()
        setContent {
            AppPrestamosTheme {
                PrestamosApp()
            }
        }
    }
}

@Composable
private fun PrestamosApp() {
    val navController = rememberNavController()
    val destinations = AppDestinations.entries

    Scaffold(
        bottomBar = {
            NavigationBar {
                val currentRoute = navController.currentBackStackEntryAsState().value
                    ?.destination
                    ?.route

                destinations.forEach { destination ->
                    NavigationBarItem(
                        selected = currentRoute == destination.route,
                        onClick = {
                            navController.navigate(destination.route) {
                                popUpTo(navController.graph.findStartDestination().id) {
                                    saveState = true
                                }
                                launchSingleTop = true
                                restoreState = true
                            }
                        },
                        icon = {
                            Icon(
                                imageVector = iconFor(destination),
                                contentDescription = destination.title
                            )
                        },
                        label = { Text(destination.title) }
                    )
                }
            }
        }
    ) { innerPadding ->
        NavHost(
            navController = navController,
            startDestination = AppDestinations.CLIENTES.route,
            modifier = Modifier.padding(innerPadding)
        ) {
            composable(AppDestinations.CLIENTES.route) { ClientesScreen() }
            composable(AppDestinations.PRESTAMOS.route) { PrestamosScreen() }
            composable(AppDestinations.PAGOS.route) { PagosScreen() }
            composable(AppDestinations.REPORTES.route) { ReportesScreen() }
        }
    }
}

@Composable
private fun iconFor(destination: AppDestinations) = when (destination) {
    AppDestinations.CLIENTES -> Icons.Outlined.Groups
    AppDestinations.PRESTAMOS -> Icons.Outlined.RequestPage
    AppDestinations.PAGOS -> Icons.Outlined.Payments
    AppDestinations.REPORTES -> Icons.Outlined.Assessment
}
