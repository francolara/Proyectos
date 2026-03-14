package com.prestamos.app.util

import java.security.MessageDigest

fun hashPin(pin: String): String {
    val digest = MessageDigest.getInstance("SHA-256")
    val bytes = digest.digest(pin.toByteArray())
    return bytes.joinToString("") { "%02x".format(it) }
}
