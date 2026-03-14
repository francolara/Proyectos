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
                message = "Se detectó manipulación de fecha. Reactiva la licencia."
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
                    message = if (isAnnualValid) "Licencia ANUAL activa" else "La licencia ANUAL venció"
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
                    message = if (trialValid) "Trial activo: quedan $daysRemaining día(s)" else "El período de prueba expiró"
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
            ?: error("Clave inválida para este equipo")

        val expiration = when (activationType) {
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
        val normalizedKey = key.trim().uppercase()
        val parts = normalizedKey.split("-")
        if (parts.size != 5) return null

        val type = parts.first()
        val keyToken = parts.drop(1).joinToString("")

        val expectedToken = when (type) {
            "ANUAL" -> buildActivationToken(deviceCode, "ANUAL")
            "FULL" -> buildActivationToken(deviceCode, "FULL")
            else -> return null
        }

        if (keyToken != expectedToken) return null
        return if (type == "ANUAL") LicenseType.ANUAL else LicenseType.FULL
    }

    fun generateActivationKey(deviceCode: String, type: LicenseType): String {
        val typeText = when (type) {
            LicenseType.ANUAL -> "ANUAL"
            LicenseType.FULL -> "FULL"
            LicenseType.TRIAL -> error("No existe clave de activación para TRIAL")
        }
        val token = buildActivationToken(deviceCode, typeText)
        return "$typeText-${token.chunked(4).joinToString("-")}"
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
