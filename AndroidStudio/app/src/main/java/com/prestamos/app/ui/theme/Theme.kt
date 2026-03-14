package com.prestamos.app.ui.theme

import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable

private val DarkColorScheme = darkColorScheme(
    primary = DarkPrimaryGreen,
    onPrimary = DarkText,
    primaryContainer = DarkSecondaryGreen,
    onPrimaryContainer = DarkText,
    secondary = DarkSecondaryGreen,
    onSecondary = DarkText,
    secondaryContainer = DarkSurface,
    onSecondaryContainer = DarkText,
    tertiary = DarkAccentGold,
    onTertiary = ColorBlack,
    background = DarkBackground,
    onBackground = DarkText,
    surface = DarkSurface,
    onSurface = DarkText
)

private val LightColorScheme = lightColorScheme(
    primary = PrimaryGreen,
    onPrimary = ColorWhite,
    primaryContainer = SecondaryGreen,
    onPrimaryContainer = AppText,
    secondary = SecondaryGreen,
    onSecondary = AppText,
    secondaryContainer = SecondaryGreen.copy(alpha = 0.22f),
    onSecondaryContainer = AppText,
    tertiary = AccentGold,
    onTertiary = AppText,
    background = AppBackground,
    onBackground = AppText,
    surface = ColorWhite,
    onSurface = AppText,
    surfaceVariant = SecondaryGreen.copy(alpha = 0.12f),
    onSurfaceVariant = AppText
)

private val ColorWhite = androidx.compose.ui.graphics.Color(0xFFFFFFFF)
private val ColorBlack = androidx.compose.ui.graphics.Color(0xFF000000)

@Composable
fun AppPrestamosTheme(
    darkTheme: Boolean = isSystemInDarkTheme(),
    dynamicColor: Boolean = false,
    content: @Composable () -> Unit
) {
    val colorScheme = if (darkTheme) DarkColorScheme else LightColorScheme

    MaterialTheme(
        colorScheme = colorScheme,
        typography = Typography,
        content = content
    )
}
