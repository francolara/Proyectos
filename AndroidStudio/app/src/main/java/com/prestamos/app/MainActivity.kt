package com.prestamos.app

import android.Manifest
import android.content.pm.PackageManager
import android.os.Build
import android.os.Bundle
import android.view.MotionEvent
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.compose.animation.core.animateDpAsState
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.fillMaxHeight
import androidx.compose.foundation.layout.fillMaxWidth
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.WindowInsets
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.statusBarsPadding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.clickable
import androidx.compose.material3.MaterialTheme
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Groups
import androidx.compose.material.icons.outlined.Home
import androidx.compose.material.icons.outlined.KeyboardArrowLeft
import androidx.compose.material.icons.outlined.KeyboardArrowRight
import androidx.compose.material.icons.outlined.Logout
import androidx.compose.material.icons.outlined.NightsStay
import androidx.compose.material.icons.outlined.Payments
import androidx.compose.material.icons.outlined.RequestPage
import androidx.compose.material.icons.outlined.Settings
import androidx.compose.material.icons.outlined.VpnKey
import androidx.compose.material3.CircularProgressIndicator
import androidx.compose.material3.Icon
import androidx.compose.material3.IconButton
import androidx.compose.material3.Text
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Surface
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.ui.ExperimentalComposeUiApi
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.input.pointer.pointerInteropFilter
import androidx.compose.ui.unit.dp
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
import kotlinx.coroutines.delay

private const val INACTIVITY_TIMEOUT_MS = 5 * 60 * 1000L

class MainActivity : ComponentActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        CuotasVencidasReminderScheduler.schedule(this)
        requestNotificationPermissionIfNeeded()
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

@OptIn(ExperimentalComposeUiApi::class)
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
    var interactionTick by remember { mutableStateOf(0) }
    val sesionActiva = authState is AuthState.Unlocked

    LaunchedEffect(authMensaje, appMensaje, activationMensaje) {
        val mensaje = authMensaje ?: appMensaje ?: activationMensaje
        if (!mensaje.isNullOrBlank()) {
            snackbarHostState.showSnackbar(mensaje)
            authViewModel.limpiarMensaje()
            appViewModel.limpiarMensaje()
            activationViewModel.limpiarMensaje()
        }
    }

    LaunchedEffect(sesionActiva) {
        if (sesionActiva) interactionTick++
    }

    LaunchedEffect(sesionActiva, interactionTick) {
        if (!sesionActiva) return@LaunchedEffect
        delay(INACTIVITY_TIMEOUT_MS)
        authViewModel.bloquearSesion()
    }

    Scaffold(
        snackbarHost = { SnackbarHost(snackbarHostState) },
        containerColor = androidx.compose.material3.MaterialTheme.colorScheme.background,
        contentWindowInsets = WindowInsets(0, 0, 0, 0)
    ) { padding ->
        Box(
            modifier = Modifier
                .fillMaxSize()
                .statusBarsPadding()
                .padding(padding)
                .pointerInteropFilter { event ->
                    if (sesionActiva && (event.actionMasked == MotionEvent.ACTION_DOWN || event.actionMasked == MotionEvent.ACTION_UP)) {
                        interactionTick++
                    }
                    false
                }
        ) {
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
                else -> PrestamosApp(appViewModel, authViewModel, activationViewModel, isDarkMode, onToggleDarkMode)
            }
        }
    }
}

@Composable
private fun PrestamosApp(
    appViewModel: AppViewModel,
    authViewModel: AuthViewModel,
    activationViewModel: ActivationViewModel,
    isDarkMode: Boolean,
    onToggleDarkMode: (Boolean) -> Unit,
    dashboardViewModel: DashboardViewModel = viewModel()
) {
    val navController = rememberNavController()
    val activationRoute = "activation_license"
    val modoNocheRoute = "__toggle_dark_mode__"
    val menuItems = buildList {
        add(SideMenuItem(modoNocheRoute, "Modo Noche"))
        AppDestinations.entries.forEach { destination ->
            if (destination == AppDestinations.LOGOUT) {
                add(SideMenuItem(activationRoute, "Licencia"))
            }
            add(SideMenuItem(destination.route, if (destination.title.isBlank()) "Salir" else destination.title))
        }
    }
    var menuExpanded by rememberSaveable { mutableStateOf(false) }
    val sideMenuWidth by animateDpAsState(targetValue = if (menuExpanded) 170.dp else 52.dp, label = "side_menu_width")

    Scaffold(
        containerColor = androidx.compose.material3.MaterialTheme.colorScheme.background,
        contentWindowInsets = WindowInsets(0, 0, 0, 0)
    ) { innerPadding ->
        Row(
            modifier = Modifier
                .fillMaxSize()
                .padding(innerPadding)
        ) {
            SideMenu(
                width = sideMenuWidth,
                expanded = menuExpanded,
                currentRoute = navController.currentBackStackEntryAsState().value?.destination?.route,
                destinations = menuItems,
                onToggleExpanded = { menuExpanded = !menuExpanded },
                onNavigate = { route ->
                    if (route == modoNocheRoute) {
                        onToggleDarkMode(!isDarkMode)
                    } else {
                        navController.navigate(route) {
                            popUpTo(navController.graph.findStartDestination().id) {
                                saveState = true
                            }
                            launchSingleTop = true
                            restoreState = true
                        }
                    }
                }
            )
            NavHost(
                navController = navController,
                startDestination = AppDestinations.DASHBOARD.route,
                modifier = Modifier
                    .weight(1f)
                    .fillMaxHeight()
                    .padding(start = 4.dp)
            ) {
                composable(AppDestinations.DASHBOARD.route) {
                    val activationState by activationViewModel.uiState.collectAsStateWithLifecycle()
                    DashboardScreen(
                        viewModel = dashboardViewModel,
                        isDarkMode = isDarkMode,
                        onToggleDarkMode = onToggleDarkMode,
                        activationUiState = activationState,
                        onActivationKeyChanged = activationViewModel::onActivationKeyChanged,
                        onActivateLicense = activationViewModel::activate,
                        onRefreshLicenseStatus = activationViewModel::refreshStatus
                    )
                }
                composable(AppDestinations.CLIENTES.route) { ClientesScreen(appViewModel) }
                composable(AppDestinations.PRESTAMOS.route) { PrestamosScreen(appViewModel) }
                composable(AppDestinations.PAGOS.route) { PagosScreen(appViewModel) }
                composable(AppDestinations.BACKUP.route) {
                    BackupScreen()
                }
                composable(activationRoute) {
                    val activationState by activationViewModel.uiState.collectAsStateWithLifecycle()
                    ActivationScreen(
                        uiState = activationState,
                        onActivationKeyChanged = activationViewModel::onActivationKeyChanged,
                        onActivate = activationViewModel::activate,
                        onRefresh = activationViewModel::refreshStatus
                    )
                }
                composable(AppDestinations.LOGOUT.route) {
                    LogoutScreen(onLogout = { authViewModel.bloquearSesion() })
                }
            }
        }
    }
}

@Composable
private fun SideMenu(
    width: androidx.compose.ui.unit.Dp,
    expanded: Boolean,
    currentRoute: String?,
    destinations: List<SideMenuItem>,
    onToggleExpanded: () -> Unit,
    onNavigate: (String) -> Unit
) {
    Surface(
        modifier = Modifier
            .fillMaxHeight()
            .width(width)
            .padding(end = 6.dp),
        color = MaterialTheme.colorScheme.primaryContainer,
        shape = RoundedCornerShape(topEnd = 16.dp, bottomEnd = 16.dp)
    ) {
        Column(
            modifier = Modifier
                .fillMaxHeight()
                .padding(horizontal = if (expanded) 8.dp else 4.dp, vertical = 8.dp)
        ) {
            IconButton(onClick = onToggleExpanded, modifier = Modifier.fillMaxWidth()) {
                Icon(
                    imageVector = if (expanded) Icons.Outlined.KeyboardArrowLeft else Icons.Outlined.KeyboardArrowRight,
                    contentDescription = "Expandir menu"
                )
            }
            if (expanded) {
                Text(
                    text = "Control de Creditos",
                    style = MaterialTheme.typography.titleSmall,
                    modifier = Modifier.padding(horizontal = 10.dp, vertical = 2.dp)
                )
            }
            Spacer(modifier = Modifier.height(6.dp))
            destinations.forEach { destination ->
                val selected = currentRoute == destination.route
                Surface(
                    color = if (selected) MaterialTheme.colorScheme.primary else Color.Transparent,
                    contentColor = if (selected) MaterialTheme.colorScheme.onPrimary else MaterialTheme.colorScheme.onPrimaryContainer,
                    shape = RoundedCornerShape(12.dp),
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(vertical = 4.dp)
                        .clickable { onNavigate(destination.route) }
                ) {
                    Row(
                        verticalAlignment = Alignment.CenterVertically,
                        modifier = Modifier.padding(horizontal = 10.dp, vertical = 10.dp)
                    ) {
                        Icon(
                            imageVector = iconForRoute(destination.route),
                            contentDescription = destination.route
                        )
                        if (expanded) {
                            Spacer(modifier = Modifier.width(12.dp))
                            Text(destination.title)
                        }
                    }
                }
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

private data class SideMenuItem(
    val route: String,
    val title: String
)

@Composable
private fun iconForRoute(route: String) = when (route) {
    "__toggle_dark_mode__" -> Icons.Outlined.NightsStay
    AppDestinations.DASHBOARD.route -> Icons.Outlined.Home
    AppDestinations.CLIENTES.route -> Icons.Outlined.Groups
    AppDestinations.PRESTAMOS.route -> Icons.Outlined.RequestPage
    AppDestinations.PAGOS.route -> Icons.Outlined.Payments
    AppDestinations.BACKUP.route -> Icons.Outlined.Settings
    AppDestinations.LOGOUT.route -> Icons.Outlined.Logout
    "activation_license" -> Icons.Outlined.VpnKey
    else -> Icons.Outlined.Home
}
