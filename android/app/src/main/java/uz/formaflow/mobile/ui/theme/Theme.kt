package uz.formaflow.mobile.ui.theme

import androidx.compose.foundation.isSystemInDarkTheme
import androidx.compose.material3.MaterialTheme
import androidx.compose.material3.darkColorScheme
import androidx.compose.material3.lightColorScheme
import androidx.compose.runtime.Composable
import androidx.compose.ui.graphics.Color

private val LightColors = lightColorScheme(
    primary = Color(0xFF1769AA),
    onPrimary = Color.White,
    primaryContainer = Color(0xFFD4E9FF),
    onPrimaryContainer = Color(0xFF001D36),
    secondary = Color(0xFF006B5E),
    secondaryContainer = Color(0xFF78F8DF),
    background = Color(0xFFF7FAFC),
    surface = Color(0xFFFFFFFF),
    surfaceVariant = Color(0xFFE8EEF3),
    error = Color(0xFFBA1A1A),
)

private val DarkColors = darkColorScheme(
    primary = Color(0xFFA3CDFF),
    onPrimary = Color(0xFF003258),
    primaryContainer = Color(0xFF004A79),
    secondary = Color(0xFF59DBC4),
    background = Color(0xFF0D141C),
    surface = Color(0xFF121B24),
    surfaceVariant = Color(0xFF27333E),
    error = Color(0xFFFFB4AB),
)

@Composable
fun FormaFlowTheme(content: @Composable () -> Unit) {
    MaterialTheme(
        colorScheme = if (isSystemInDarkTheme()) DarkColors else LightColors,
        content = content,
    )
}
