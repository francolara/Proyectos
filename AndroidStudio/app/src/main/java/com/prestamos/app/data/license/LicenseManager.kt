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
        val parts = key.split("-")
        if (parts.size != 3) return null
        val type = parts[0]
        val keyDevice = parts[1]
        val signature = parts[2]
        if (keyDevice != deviceCode) return null

        val expected = when (type) {
            "ANUAL" -> buildSignature("ANUAL", deviceCode)
            "FULL" -> buildSignature("FULL", deviceCode)
            else -> return null
        }

        if (signature != expected) return null
        return if (type == "ANUAL") LicenseType.ANUAL else LicenseType.FULL
    }

    private fun buildSignature(type: String, deviceCode: String): String {
        return sha256("$type|$deviceCode|$LICENSE_SECRET").take(16).uppercase()
    }

    private fun sha256(value: String): String {
        val digest = MessageDigest.getInstance("SHA-256")
        return digest.digest(value.toByteArray()).joinToString("") { "%02x".format(it) }
    }

    companion object {
        private const val TRIAL_DAYS = 30L
        private const val LICENSE_SECRET = "PRESTAMOS_APP_OFFLINE_LICENSE_v1"
    }
}
