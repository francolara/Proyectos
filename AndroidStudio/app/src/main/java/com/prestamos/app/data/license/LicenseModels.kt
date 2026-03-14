package com.prestamos.app.data.license

enum class LicenseType {
    TRIAL,
    ANUAL,
    FULL
}

data class LicenseStatus(
    val isValid: Boolean = false,
    val deviceCode: String = "",
    val licenseType: LicenseType = LicenseType.TRIAL,
    val isActivated: Boolean = false,
    val trialDaysRemaining: Long = 0,
    val trialExpired: Boolean = false,
    val activationDate: Long? = null,
    val expirationDate: Long? = null,
    val manipulatedDateDetected: Boolean = false,
    val message: String = ""
)
