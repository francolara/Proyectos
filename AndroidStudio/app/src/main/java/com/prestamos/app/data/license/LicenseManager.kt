package com.prestamos.app.data.license

import android.content.Context
import android.provider.Settings
import kotlinx.coroutines.flow.first
import java.security.MessageDigest
import java.util.concurrent.TimeUnit

class LicenseManager(private val context: Context) {
    private val prefs = LicensePreferences(context)

    suspend fun evaluateStatus(): LicenseStatus {
        val now = System.currentTimeMillis()
        val deviceCode = generateDeviceCode()
        val trialEnd = now + TimeUnit.DAYS.toMillis(TRIAL_DAYS)

        prefs.initializeIfNeeded(
            deviceCode = deviceCode,
            now = now,
            trialEnd = trialEnd
        )

        val data = prefs.rawData.first()
        val storedDeviceCode = (data["deviceCode"] as? String).orEmpty().ifBlank { deviceCode }
        val storedType = (data["licenseType"] as? String)
            ?.let { runCatching { LicenseType.valueOf(it) }.getOrNull() }
            ?: LicenseType.TRIAL
        val isActivated = data["isActivated"] as? Boolean ?: false
        val trialEndDate = data["trialEndDate"] as? Long ?: trialEnd
        val activationDate = data["activationDate"] as? Long
        val expirationDate = data["expirationDate"] as? Long
        val lastValidUseDate = data["lastValidUseDate"] as? Long ?: now

        if (now < lastValidUseDate) {
            return LicenseStatus(
                isValid = false,
                deviceCode = storedDeviceCode,
                licenseType = storedType,
                isActivated = isActivated,
                trialDaysRemaining = 0,
                trialExpired = true,
                activationDate = activationDate,
                expirationDate = expirationDate,
                manipulatedDateDetected = true,
                message = "Se detecto manipulacion de fecha. Reactiva la licencia."
            )
        }

        val status = when {
            isActivated && storedType == LicenseType.FULL -> {
                LicenseStatus(
                    isValid = true,
                    deviceCode = storedDeviceCode,
                    licenseType = LicenseType.FULL,
                    isActivated = true,
                    trialDaysRemaining = 0,
                    trialExpired = false,
                    activationDate = activationDate,
                    expirationDate = null,
                    manipulatedDateDetected = false,
                    message = "Licencia FULL activa"
                )
            }

            isActivated && storedType == LicenseType.ANUAL -> {
                val isAnnualValid = expirationDate != null && now <= expirationDate
                LicenseStatus(
                    isValid = isAnnualValid,
                    deviceCode = storedDeviceCode,
                    licenseType = LicenseType.ANUAL,
                    isActivated = true,
                    trialDaysRemaining = 0,
                    trialExpired = !isAnnualValid,
                    activationDate = activationDate,
                    expirationDate = expirationDate,
                    manipulatedDateDetected = false,
                    message = if (isAnnualValid) "Licencia ANUAL activa" else "La licencia ANUAL vencio"
                )
            }

            isActivated && storedType == LicenseType.MENSUAL -> {
                val isMonthlyValid = expirationDate != null && now <= expirationDate
                LicenseStatus(
                    isValid = isMonthlyValid,
                    deviceCode = storedDeviceCode,
                    licenseType = LicenseType.MENSUAL,
                    isActivated = true,
                    trialDaysRemaining = 0,
                    trialExpired = !isMonthlyValid,
                    activationDate = activationDate,
                    expirationDate = expirationDate,
                    manipulatedDateDetected = false,
                    message = if (isMonthlyValid) "Licencia MENSUAL activa" else "La licencia MENSUAL vencio"
                )
            }

            else -> {
                val remainingMillis = trialEndDate - now
                val daysRemaining = if (remainingMillis <= 0L) 0L else TimeUnit.MILLISECONDS.toDays(remainingMillis).coerceAtLeast(1)
                val trialValid = now <= trialEndDate
                LicenseStatus(
                    isValid = trialValid,
                    deviceCode = storedDeviceCode,
                    licenseType = LicenseType.TRIAL,
                    isActivated = false,
                    trialDaysRemaining = daysRemaining,
                    trialExpired = !trialValid,
                    activationDate = activationDate,
                    expirationDate = expirationDate,
                    manipulatedDateDetected = false,
                    message = if (trialValid) "Trial activo: quedan $daysRemaining dia(s)" else "El periodo de prueba expiro"
                )
            }
        }

        if (status.isValid) {
            prefs.saveLastValidUseDate(now)
        }

        return status
    }

    suspend fun activate(key: String): LicenseStatus {
        val now = System.currentTimeMillis()
        val status = evaluateStatus()
        val normalized = key.trim().uppercase()
        val activationType = validateActivationKey(status.deviceCode, normalized)
            ?: error("Clave invalida para este equipo")

        val expiration = when (activationType) {
            LicenseType.MENSUAL -> now + TimeUnit.DAYS.toMillis(30)
            LicenseType.ANUAL -> now + TimeUnit.DAYS.toMillis(365)
            LicenseType.FULL -> null
            LicenseType.TRIAL -> null
        }

        prefs.saveActivation(
            type = activationType,
            licenseKey = normalized,
            activationDate = now,
            expirationDate = expiration
        )

        return evaluateStatus()
    }

    fun generateDeviceCode(): String {
        val androidId = Settings.Secure.getString(context.contentResolver, Settings.Secure.ANDROID_ID).orEmpty()
        val installTime = runCatching {
            context.packageManager.getPackageInfo(context.packageName, 0).firstInstallTime
        }.getOrDefault(0L)
        return sha256("$androidId|${context.packageName}|$installTime").take(12).uppercase()
    }

    fun validateActivationKey(deviceCode: String, key: String): LicenseType? {
        val normalizedKey = key.trim().uppercase().replace(" ", "")

        // Compatibilidad: acepta formato antiguo con prefijo y formato nuevo solo codigo.
        val prefixedType = when {
            normalizedKey.startsWith("MENSUAL-") -> LicenseType.MENSUAL
            normalizedKey.startsWith("ANUAL-") -> LicenseType.ANUAL
            normalizedKey.startsWith("FULL-") -> LicenseType.FULL
            else -> null
        }

        val keyToken = if (prefixedType != null) {
            normalizedKey.substringAfter("-").replace("-", "")
        } else {
            normalizedKey.replace("-", "")
        }

        if (keyToken.length != 16) return null

        val expectedByType = listOf(
            LicenseType.MENSUAL to buildActivationToken(deviceCode, "MENSUAL"),
            LicenseType.ANUAL to buildActivationToken(deviceCode, "ANUAL"),
            LicenseType.FULL to buildActivationToken(deviceCode, "FULL")
        )

        if (prefixedType != null) {
            val expected = expectedByType.firstOrNull { it.first == prefixedType }?.second ?: return null
            return if (keyToken == expected) prefixedType else null
        }

        return expectedByType.firstOrNull { it.second == keyToken }?.first
    }

    fun generateActivationKey(deviceCode: String, type: LicenseType): String {
        val typeText = when (type) {
            LicenseType.MENSUAL -> "MENSUAL"
            LicenseType.ANUAL -> "ANUAL"
            LicenseType.FULL -> "FULL"
            LicenseType.TRIAL -> error("No existe clave de activacion para TRIAL")
        }
        val token = buildActivationToken(deviceCode, typeText)
        return token.chunked(4).joinToString("-")
    }

    private fun buildActivationToken(deviceCode: String, typeText: String): String {
        return sha256("$deviceCode$typeText$LICENSE_SECRET").take(16).uppercase()
    }

    private fun sha256(value: String): String {
        val digest = MessageDigest.getInstance("SHA-256")
        return digest.digest(value.toByteArray()).joinToString("") { "%02x".format(it) }
    }

    companion object {
        private const val TRIAL_DAYS = 30L
        private const val LICENSE_SECRET = "PRST_APP_LICENSE_PRIVATE_2026_#84KF92@A"
    }
}
