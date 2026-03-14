package com.prestamos.app

import android.Manifest
import android.content.pm.PackageManager
import android.os.Build
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
import androidx.compose.material.icons.outlined.Logout
import androidx.compose.material.icons.outlined.Payments
import androidx.compose.material.icons.outlined.RequestPage
import androidx.compose.material.icons.outlined.Settings
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.Icon
import androidx.compose.material3.NavigationBar
import androidx.compose.material3.NavigationBarItem
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.core.content.ContextCompat
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.navigation.NavGraph.Companion.findStartDestination
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.currentBackStackEntryAsState
import androidx.navigation.compose.rememberNavController
import com.prestamos.app.navigation.AppDestinations
import com.prestamos.app.notifications.CuotasVencidasReminderScheduler
import com.prestamos.app.ui.screen.ClientesScreen
import com.prestamos.app.ui.screen.DashboardScreen
import com.prestamos.app.ui.screen.LogoutScreen
import com.prestamos.app.ui.screen.PagosScreen
import com.prestamos.app.ui.screen.BackupScreen
import com.prestamos.app.ui.screen.ActivationScreen
import com.prestamos.app.ui.screen.PinLoginScreen
import com.prestamos.app.ui.screen.PinSetupScreen
import com.prestamos.app.ui.screen.PrestamosScreen
import com.prestamos.app.ui.theme.AppPrestamosTheme
import com.prestamos.app.ui.viewmodel.ActivationViewModel
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.ui.viewmodel.AuthState
import com.prestamos.app.ui.viewmodel.AuthViewModel
import com.prestamos.app.ui.viewmodel.DashboardViewModel

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        CuotasVencidasReminderScheduler.schedule(this)
        requestNotificationPermissionIfNeeded()
        enableEdgeToEdge()
        setContent {
            var darkMode by rememberSaveable { mutableStateOf(false) }
            AppPrestamosTheme(darkTheme = darkMode) {
                AppRoot(
                    isDarkMode = darkMode,
                    onToggleDarkMode = { darkMode = it }
                )
            }
        }
    }

    private fun requestNotificationPermissionIfNeeded() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.TIRAMISU) return
        val granted = ContextCompat.checkSelfPermission(this, Manifest.permission.POST_NOTIFICATIONS) == PackageManager.PERMISSION_GRANTED
        if (!granted) {
            requestPermissions(arrayOf(Manifest.permission.POST_NOTIFICATIONS), 3001)
        }
    }
}

@Composable
private fun AppRoot(
    isDarkMode: Boolean,
    onToggleDarkMode: (Boolean) -> Unit,
    authViewModel: AuthViewModel = viewModel(),
    appViewModel: AppViewModel = viewModel(),
    activationViewModel: ActivationViewModel = viewModel()
) {
    val authState by authViewModel.authState.collectAsStateWithLifecycle()
    val authMensaje by authViewModel.mensaje.collectAsStateWithLifecycle()
    val appMensaje by appViewModel.mensaje.collectAsStateWithLifecycle()
    val activationState by activationViewModel.uiState.collectAsStateWithLifecycle()
    val activationMensaje by activationViewModel.mensaje.collectAsStateWithLifecycle()

    val snackbarHostState = remember { SnackbarHostState() }

    LaunchedEffect(authMensaje, appMensaje, activationMensaje) {
        val mensaje = authMensaje ?: appMensaje ?: activationMensaje
        if (!mensaje.isNullOrBlank()) {
            snackbarHostState.showSnackbar(mensaje)
            authViewModel.limpiarMensaje()
            appViewModel.limpiarMensaje()
            activationViewModel.limpiarMensaje()
        }
    }

    Scaffold(snackbarHost = { SnackbarHost(snackbarHostState) }) { padding ->
        Box(modifier = Modifier.fillMaxSize().padding(padding)) {
            when {
                activationState.loading || authState is AuthState.Loading -> Box(
                    contentAlignment = Alignment.Center,
                    modifier = Modifier.fillMaxSize()
                ) {
                    CircularProgressIndicator()
                }

                !activationState.canAccessApp -> ActivationScreen(
                    uiState = activationState,
                    onActivationKeyChanged = activationViewModel::onActivationKeyChanged,
                    onActivate = activationViewModel::activate,
                    onRefresh = activationViewModel::refreshStatus
                )

                authState is AuthState.NeedsPinSetup -> PinSetupScreen(authViewModel)
                authState is AuthState.Locked -> PinLoginScreen(authViewModel)
                else -> PrestamosApp(appViewModel, authViewModel, isDarkMode, onToggleDarkMode)
            }
        }
    }
}

@Composable
private fun PrestamosApp(
    appViewModel: AppViewModel,
    authViewModel: AuthViewModel,
    isDarkMode: Boolean,
    onToggleDarkMode: (Boolean) -> Unit,
    dashboardViewModel: DashboardViewModel = viewModel()
) {
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
                                contentDescription = if (destination.title.isBlank()) destination.route else destination.title
                            )
                        },
                        label = null
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
                    isDarkMode = isDarkMode,
                    onToggleDarkMode = onToggleDarkMode
                )
            }
            composable(AppDestinations.CLIENTES.route) { ClientesScreen(appViewModel) }
            composable(AppDestinations.PRESTAMOS.route) { PrestamosScreen(appViewModel) }
            composable(AppDestinations.PAGOS.route) { PagosScreen(appViewModel) }
            composable(AppDestinations.BACKUP.route) { BackupScreen() }
            composable(AppDestinations.LOGOUT.route) {
                LogoutScreen(onLogout = { authViewModel.bloquearSesion() })
            }
        }
    }
}

@Composable
private fun iconFor(destination: AppDestinations) = when (destination) {
    AppDestinations.DASHBOARD -> Icons.Outlined.Home
    AppDestinations.CLIENTES -> Icons.Outlined.Groups
    AppDestinations.PRESTAMOS -> Icons.Outlined.RequestPage
    AppDestinations.PAGOS -> Icons.Outlined.Payments
    AppDestinations.BACKUP -> Icons.Outlined.Settings
    AppDestinations.LOGOUT -> Icons.Outlined.Logout
}
