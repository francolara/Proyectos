package com.prestamos.app.data.security

import android.content.Context
import com.prestamos.app.util.hashPin
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.first

class AuthRepository(context: Context) {
    private val prefs = AuthPreferences(context)

    val pinConfigured: Flow<Boolean> = prefs.pinConfigured
    val sessionUnlocked: Flow<Boolean> = prefs.sessionUnlocked

    suspend fun configurarPin(pin: String) {
        prefs.savePinHash(hashPin(pin))
        prefs.setSessionUnlocked(true)
    }

    suspend fun validarPin(pin: String): Boolean {
        val hash = prefs.pinHash.first()
        val ok = hash == hashPin(pin)
        if (ok) {
            prefs.setSessionUnlocked(true)
        }
        return ok
    }

    suspend fun desbloquearSesion() {
        prefs.setSessionUnlocked(true)
    }

    suspend fun bloquearSesion() {
        prefs.setSessionUnlocked(false)
    }
}
