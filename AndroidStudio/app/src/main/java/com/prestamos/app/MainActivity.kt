package com.prestamos.app

import android.os.Bundle
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.enableEdgeToEdge
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Groups
import androidx.compose.material.icons.outlined.Home
import androidx.compose.material.icons.outlined.Payments
import androidx.compose.material.icons.outlined.RequestPage
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.Icon
import androidx.compose.material3.NavigationBar
import androidx.compose.material3.NavigationBarItem
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Text
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.remember
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.navigation.NavGraph.Companion.findStartDestination
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.currentBackStackEntryAsState
import androidx.navigation.compose.rememberNavController
import com.prestamos.app.navigation.AppDestinations
import com.prestamos.app.ui.screen.ClientesScreen
import com.prestamos.app.ui.screen.DashboardScreen
import com.prestamos.app.ui.screen.PagosScreen
import com.prestamos.app.ui.screen.PinLoginScreen
import com.prestamos.app.ui.screen.PinSetupScreen
import com.prestamos.app.ui.screen.PrestamosScreen
import com.prestamos.app.ui.theme.AppPrestamosTheme
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.ui.viewmodel.AuthState
import com.prestamos.app.ui.viewmodel.AuthViewModel
import com.prestamos.app.ui.viewmodel.DashboardViewModel

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
        setContent {
            AppPrestamosTheme {
                AppRoot()
            }
        }
    }
}

@Composable
private fun AppRoot(
    authViewModel: AuthViewModel = viewModel(),
    appViewModel: AppViewModel = viewModel()
) {
    val authState by authViewModel.authState.collectAsStateWithLifecycle()
    val authMensaje by authViewModel.mensaje.collectAsStateWithLifecycle()
    val appMensaje by appViewModel.mensaje.collectAsStateWithLifecycle()

    val snackbarHostState = remember { SnackbarHostState() }

    LaunchedEffect(authMensaje, appMensaje) {
        val mensaje = authMensaje ?: appMensaje
        if (!mensaje.isNullOrBlank()) {
            snackbarHostState.showSnackbar(mensaje)
            authViewModel.limpiarMensaje()
            appViewModel.limpiarMensaje()
        }
    }

    Scaffold(snackbarHost = { SnackbarHost(snackbarHostState) }) { padding ->
        Box(modifier = Modifier.fillMaxSize().padding(padding)) {
            when (authState) {
                AuthState.Loading -> Box(contentAlignment = Alignment.Center, modifier = Modifier.fillMaxSize()) {
                    CircularProgressIndicator()
                }

                AuthState.NeedsPinSetup -> PinSetupScreen(authViewModel)
                AuthState.Locked -> PinLoginScreen(authViewModel)
                AuthState.Unlocked -> PrestamosApp(appViewModel, authViewModel)
            }
        }
    }
}

@Composable
private fun PrestamosApp(appViewModel: AppViewModel, authViewModel: AuthViewModel, dashboardViewModel: DashboardViewModel = viewModel()) {
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
            startDestination = AppDestinations.DASHBOARD.route,
            modifier = Modifier.padding(innerPadding)
        ) {
            composable(AppDestinations.DASHBOARD.route) {
                DashboardScreen(
                    viewModel = dashboardViewModel,
                    onLogout = { authViewModel.bloquearSesion() },
                    onNavigate = { route -> navController.navigate(route) }
                )
            }
            composable(AppDestinations.CLIENTES.route) { ClientesScreen(appViewModel) }
            composable(AppDestinations.PRESTAMOS.route) { PrestamosScreen(appViewModel) }
            composable(AppDestinations.PAGOS.route) { PagosScreen(appViewModel) }
        }
    }
}

@Composable
private fun iconFor(destination: AppDestinations) = when (destination) {
    AppDestinations.DASHBOARD -> Icons.Outlined.Home
    AppDestinations.CLIENTES -> Icons.Outlined.Groups
    AppDestinations.PRESTAMOS -> Icons.Outlined.RequestPage
    AppDestinations.PAGOS -> Icons.Outlined.Payments
}
