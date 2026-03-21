package com.prestamos.app

import android.Manifest
import android.app.Activity
import android.app.KeyguardManager
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
import androidx.compose.foundation.layout.size
import androidx.compose.foundation.layout.WindowInsets
import androidx.compose.foundation.layout.padding
import androidx.compose.foundation.layout.navigationBarsPadding
import androidx.compose.foundation.layout.Row
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.statusBarsPadding
import androidx.compose.foundation.layout.width
import androidx.compose.foundation.clickable
import androidx.compose.foundation.background
import androidx.compose.material3.MaterialTheme
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.outlined.Groups
import androidx.compose.material.icons.outlined.Inventory2
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
import androidx.compose.material3.HorizontalDivider
import androidx.compose.material3.Text
import androidx.compose.material3.Scaffold
import androidx.compose.material3.SnackbarHost
import androidx.compose.material3.SnackbarHostState
import androidx.compose.material3.Surface
import androidx.compose.runtime.Composable
import androidx.compose.runtime.DisposableEffect
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.remember
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.runtime.rememberUpdatedState
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.ui.ExperimentalComposeUiApi
import androidx.compose.ui.draw.clip
import androidx.compose.ui.graphics.Brush
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.input.pointer.pointerInput
import androidx.compose.ui.input.pointer.pointerInteropFilter
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.content.ContextCompat
import androidx.lifecycle.compose.collectAsStateWithLifecycle
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.lifecycle.Lifecycle
import androidx.lifecycle.LifecycleEventObserver
import androidx.lifecycle.compose.LocalLifecycleOwner
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
import com.prestamos.app.ui.screen.ConfiguracionScreen
import com.prestamos.app.ui.screen.OnboardingScreen
import com.prestamos.app.ui.screen.PinLoginScreen
import com.prestamos.app.ui.screen.PinSetupScreen
import com.prestamos.app.ui.screen.PrestamosScreen
import com.prestamos.app.ui.theme.AppPrestamosTheme
import com.prestamos.app.ui.viewmodel.ActivationViewModel
import com.prestamos.app.ui.viewmodel.AppViewModel
import com.prestamos.app.ui.viewmodel.AuthState
import com.prestamos.app.ui.viewmodel.AuthViewModel
import com.prestamos.app.ui.viewmodel.DashboardViewModel
import com.prestamos.app.ui.viewmodel.OnboardingViewModel
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
    onboardingViewModel: OnboardingViewModel = viewModel(),
    authViewModel: AuthViewModel = viewModel(),
    appViewModel: AppViewModel = viewModel(),
    activationViewModel: ActivationViewModel = viewModel()
) {
    val onboardingState by onboardingViewModel.uiState.collectAsStateWithLifecycle()
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

    val sesionActivaActual by rememberUpdatedState(sesionActiva)
    val lifecycleOwner = LocalLifecycleOwner.current
    val context = LocalContext.current
    val activity = context as? Activity
    val keyguardManager = context.getSystemService(KeyguardManager::class.java)
    DisposableEffect(lifecycleOwner) {
        val observer = LifecycleEventObserver { _, event ->
            if (event == Lifecycle.Event.ON_STOP &&
                sesionActivaActual &&
                activity?.isChangingConfigurations != true &&
                keyguardManager?.isDeviceLocked == true
            ) {
                authViewModel.bloquearSesion()
            }
        }
        lifecycleOwner.lifecycle.addObserver(observer)
        onDispose {
            lifecycleOwner.lifecycle.removeObserver(observer)
        }
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
                .navigationBarsPadding()
                .padding(padding)
                .pointerInteropFilter { event ->
                    if (sesionActiva && (event.actionMasked == MotionEvent.ACTION_DOWN || event.actionMasked == MotionEvent.ACTION_UP)) {
                        interactionTick++
                    }
                    false
                }
        ) {
            when {
                onboardingState.loading || activationState.loading || authState is AuthState.Loading -> Box(
                    contentAlignment = Alignment.Center,
                    modifier = Modifier.fillMaxSize()
                ) {
                    CircularProgressIndicator()
                }

                onboardingState.showOnboarding -> OnboardingScreen(
                    uiState = onboardingState,
                    onComenzar = onboardingViewModel::comenzar,
                    onBusinessNameChange = onboardingViewModel::updateBusinessName,
                    onMainCurrencySelected = onboardingViewModel::selectMainCurrency,
                    onSecondaryCurrencySelected = onboardingViewModel::selectSecondaryCurrency,
                    onFinalizar = onboardingViewModel::finalizarConfiguracion
                )

                !activationState.canAccessApp -> ActivationScreen(
                    uiState = activationState,
                    onActivationKeyChanged = activationViewModel::onActivationKeyChanged,
                    onActivate = activationViewModel::activate
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
    val configuracionRoute = "settings_configuracion"
    val modoNocheRoute = "__toggle_dark_mode__"
    val menuItems = buildList {
        add(SideMenuItem(modoNocheRoute, "Modo Noche"))
        AppDestinations.entries.forEach { destination ->
            if (destination == AppDestinations.LOGOUT) {
                add(SideMenuItem(activationRoute, "Licencia"))
                add(SideMenuItem(configuracionRoute, "Configuracion"))
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
                darkModeRoute = modoNocheRoute,
                isDarkMode = isDarkMode,
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
                    menuExpanded = false
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
                        onActivate = activationViewModel::activate
                    )
                }
                composable(configuracionRoute) {
                    ConfiguracionScreen(appViewModel)
                }
                composable(AppDestinations.LOGOUT.route) {
                    LogoutScreen(onLogout = { authViewModel.cerrarSesionConRespaldo() })
                }
            }
        }
    }
}

@Composable
private fun SideMenu(
    width: androidx.compose.ui.unit.Dp,
    expanded: Boolean,
    darkModeRoute: String,
    isDarkMode: Boolean,
    currentRoute: String?,
    destinations: List<SideMenuItem>,
    onToggleExpanded: () -> Unit,
    onNavigate: (String) -> Unit
) {
    Box(
        modifier = Modifier
            .fillMaxHeight()
            .width(width)
            .padding(end = 6.dp)
            .clip(RoundedCornerShape(topEnd = 18.dp, bottomEnd = 18.dp))
            .background(
                brush = Brush.verticalGradient(
                    colors = listOf(Color(0xFF66BB6A), Color(0xFF1B5E20))
                )
            )
    ) {
        Column(
            modifier = Modifier
                .fillMaxHeight()
                .fillMaxWidth()
                .padding(horizontal = if (expanded) 10.dp else 5.dp, vertical = 10.dp)
        ) {
            Box(modifier = Modifier.fillMaxWidth()) {
                Surface(
                    color = Color.White.copy(alpha = 0.16f),
                    shape = RoundedCornerShape(12.dp),
                    modifier = Modifier
                        .align(if (expanded) Alignment.CenterEnd else Alignment.Center)
                        .size(36.dp)
                ) {
                    IconButton(onClick = onToggleExpanded) {
                        Icon(
                            imageVector = if (expanded) Icons.Outlined.KeyboardArrowLeft else Icons.Outlined.KeyboardArrowRight,
                            contentDescription = "Expandir menu",
                            tint = Color.White
                        )
                    }
                }
            }
            if (expanded) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 8.dp, vertical = 8.dp)
                ) {
                    Text(
                        text = "Control de Creditos",
                        style = MaterialTheme.typography.titleMedium.copy(
                            fontWeight = FontWeight.SemiBold,
                            letterSpacing = 0.3.sp
                        ),
                        color = Color(0xFFF5FFF5)
                    )
                    Text(
                        text = "CrediControl",
                        style = MaterialTheme.typography.bodySmall,
                        color = Color.White.copy(alpha = 0.88f)
                    )
                }
            }
            Spacer(modifier = Modifier.height(if (expanded) 8.dp else 6.dp))
            Column(
                modifier = Modifier
                    .weight(1f)
                    .fillMaxWidth()
            ) {
                destinations.forEach { destination ->
                    val isModeNight = destination.route == darkModeRoute
                    val selected = if (isModeNight) isDarkMode else currentRoute == destination.route
                    val itemColor = when {
                        selected -> Color.White.copy(alpha = 0.24f)
                        isModeNight -> Color.White.copy(alpha = 0.10f)
                        else -> Color.Transparent
                    }
                    Surface(
                        color = itemColor,
                        contentColor = Color(0xFFF7FFF7),
                        shape = RoundedCornerShape(13.dp),
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 5.dp)
                            .clickable { onNavigate(destination.route) }
                    ) {
                        Row(
                            verticalAlignment = Alignment.CenterVertically,
                            modifier = Modifier.padding(
                                horizontal = if (expanded) 12.dp else 8.dp,
                                vertical = if (expanded) 11.dp else 10.dp
                            )
                        ) {
                            Box(
                                modifier = Modifier
                                    .width(3.dp)
                                    .height(22.dp)
                                    .background(
                                        color = if (selected) Color.White else Color.Transparent,
                                        shape = RoundedCornerShape(50)
                                    )
                            )
                            Spacer(modifier = Modifier.width(8.dp))
                            Icon(
                                imageVector = iconForRoute(destination.route),
                                contentDescription = destination.route,
                                tint = Color(0xFFF7FFF7)
                            )
                            if (expanded) {
                                Spacer(modifier = Modifier.width(12.dp))
                                Text(
                                    destination.title,
                                    maxLines = 1,
                                    overflow = TextOverflow.Ellipsis,
                                    style = MaterialTheme.typography.bodyMedium.copy(
                                        fontWeight = if (selected) FontWeight.SemiBold else FontWeight.Medium
                                    ),
                                    color = Color(0xFFF7FFF7)
                                )
                            }
                        }
                    }
                }
            }
            if (expanded) {
                Column(
                    modifier = Modifier
                        .fillMaxWidth()
                        .padding(horizontal = 8.dp, vertical = 4.dp)
                ) {
                    HorizontalDivider(color = Color.White.copy(alpha = 0.24f))
                    Spacer(modifier = Modifier.height(8.dp))
                    Text(
                        text = "Contacto",
                        style = MaterialTheme.typography.labelSmall,
                        color = Color.White.copy(alpha = 0.90f)
                    )
                    Text(
                        text = "franko.laras@gmail.com",
                        style = MaterialTheme.typography.bodySmall.copy(fontWeight = FontWeight.SemiBold),
                        maxLines = 1,
                        overflow = TextOverflow.Ellipsis,
                        color = Color.White.copy(alpha = 0.96f)
                    )
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
    AppDestinations.BACKUP -> Icons.Outlined.Inventory2
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
    AppDestinations.BACKUP.route -> Icons.Outlined.Inventory2
    AppDestinations.LOGOUT.route -> Icons.Outlined.Logout
    "activation_license" -> Icons.Outlined.VpnKey
    "settings_configuracion" -> Icons.Outlined.Settings
    else -> Icons.Outlined.Home
}
