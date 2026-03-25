package com.prestamos.app.data.backup

import android.security.keystore.KeyGenParameterSpec
import android.security.keystore.KeyProperties
import java.nio.ByteBuffer
import java.security.KeyStore
import java.security.MessageDigest
import java.security.SecureRandom
import javax.crypto.AEADBadTagException
import javax.crypto.Cipher
import javax.crypto.KeyGenerator
import javax.crypto.SecretKeyFactory
import javax.crypto.SecretKey
import javax.crypto.spec.GCMParameterSpec
import javax.crypto.spec.PBEKeySpec
import javax.crypto.spec.SecretKeySpec

object BackupCryptoManager {
    private const val ANDROID_KEYSTORE = "AndroidKeyStore"
    private const val LEGACY_KEY_ALIAS = "app_prestamos_backup_key_v1"
    private const val TRANSFORMATION = "AES/GCM/NoPadding"
    private const val IV_LENGTH_BYTES = 12
    private const val GCM_TAG_BITS = 128
    private const val SALT_LENGTH_BYTES = 16
    private const val KEY_SIZE_BITS = 256
    private const val PBKDF2_ITERATIONS = 120_000
    private const val HASH_LENGTH_BYTES = 32
    private val MAGIC_PORTABLE = byteArrayOf('A'.code.toByte(), 'P'.code.toByte(), 'B'.code.toByte(), 'P'.code.toByte())
    private val MAGIC_LEGACY = byteArrayOf('A'.code.toByte(), 'P'.code.toByte(), 'B'.code.toByte(), 'K'.code.toByte())
    private const val VERSION_PORTABLE: Byte = 1

    fun encryptPortable(input: ByteArray, password: String): ByteArray {
        require(password.isNotBlank()) { "Configura una clave de respaldo antes de continuar" }
        val salt = ByteArray(SALT_LENGTH_BYTES).also { SecureRandom().nextBytes(it) }
        val secretKey = deriveKey(password, salt)
        val cipher = Cipher.getInstance(TRANSFORMATION)
        cipher.init(Cipher.ENCRYPT_MODE, secretKey)
        val iv = cipher.iv
        require(iv.size == IV_LENGTH_BYTES) { "No se pudo generar IV de cifrado valido" }
        val encrypted = cipher.doFinal(input)
        val hash = MessageDigest.getInstance("SHA-256").digest(input)
        val buffer = ByteBuffer.allocate(
            MAGIC_PORTABLE.size + 1 + 1 + 1 + salt.size + iv.size + HASH_LENGTH_BYTES + encrypted.size
        )
        buffer.put(MAGIC_PORTABLE)
        buffer.put(VERSION_PORTABLE)
        buffer.put(salt.size.toByte())
        buffer.put(iv.size.toByte())
        buffer.put(salt)
        buffer.put(iv)
        buffer.put(hash)
        buffer.put(encrypted)
        return buffer.array()
    }

    fun decryptAuto(input: ByteArray, password: String?): ByteArray {
        return when {
            hasHeader(input, MAGIC_PORTABLE) -> decryptPortable(input, password)
            hasHeader(input, MAGIC_LEGACY) -> decryptLegacyKeystore(input)
            else -> input
        }
    }

    private fun decryptPortable(input: ByteArray, password: String?): ByteArray {
        require(!password.isNullOrBlank()) { "Configura tu clave de respaldo para restaurar este archivo" }
        require(input.size > MAGIC_PORTABLE.size + 3) { "Backup cifrado invalido" }

        val buffer = ByteBuffer.wrap(input)
        val magic = ByteArray(MAGIC_PORTABLE.size)
        buffer.get(magic)
        require(magic.contentEquals(MAGIC_PORTABLE)) { "Formato de backup invalido" }
        val version = buffer.get()
        require(version == VERSION_PORTABLE) { "Version de backup no compatible" }

        val saltLength = buffer.get().toInt() and 0xFF
        val ivLength = buffer.get().toInt() and 0xFF
        require(saltLength in 12..32) { "Backup cifrado invalido" }
        require(ivLength in 12..16) { "Backup cifrado invalido" }
        require(buffer.remaining() > (saltLength + ivLength + HASH_LENGTH_BYTES)) { "Backup cifrado incompleto" }

        val salt = ByteArray(saltLength)
        val iv = ByteArray(ivLength)
        buffer.get(salt)
        buffer.get(iv)
        val expectedHash = ByteArray(HASH_LENGTH_BYTES)
        buffer.get(expectedHash)
        val encrypted = ByteArray(buffer.remaining())
        buffer.get(encrypted)

        val cipher = Cipher.getInstance(TRANSFORMATION)
        val secretKey = deriveKey(password, salt)
        val plain = try {
            cipher.init(Cipher.DECRYPT_MODE, secretKey, GCMParameterSpec(GCM_TAG_BITS, iv))
            cipher.doFinal(encrypted)
        } catch (_: AEADBadTagException) {
            throw IllegalStateException("Clave de respaldo incorrecta o archivo alterado")
        }
        val actualHash = MessageDigest.getInstance("SHA-256").digest(plain)
        require(expectedHash.contentEquals(actualHash)) { "El respaldo cifrado es invalido o fue alterado" }
        return plain
    }

    private fun hasHeader(input: ByteArray, magic: ByteArray): Boolean {
        if (input.size < magic.size + 2) return false
        for (i in magic.indices) {
            if (input[i] != magic[i]) return false
        }
        return true
    }

    private fun decryptLegacyKeystore(input: ByteArray): ByteArray {
        require(input.size > MAGIC_LEGACY.size + 2) { "Backup cifrado invalido" }
        val buffer = ByteBuffer.wrap(input)
        val magic = ByteArray(MAGIC_LEGACY.size)
        buffer.get(magic)
        require(magic.contentEquals(MAGIC_LEGACY)) { "Formato de backup invalido" }
        val version = buffer.get()
        require(version == VERSION_PORTABLE) { "Version de backup no compatible" }
        val ivLength = buffer.get().toInt() and 0xFF
        require(ivLength in 12..16) { "Backup cifrado invalido" }
        require(buffer.remaining() > ivLength) { "Backup cifrado incompleto" }
        val iv = ByteArray(ivLength)
        buffer.get(iv)
        val encrypted = ByteArray(buffer.remaining())
        buffer.get(encrypted)

        val cipher = Cipher.getInstance(TRANSFORMATION)
        cipher.init(Cipher.DECRYPT_MODE, getOrCreateLegacySecretKey(), GCMParameterSpec(GCM_TAG_BITS, iv))
        return cipher.doFinal(encrypted)
    }

    private fun deriveKey(password: String, salt: ByteArray): SecretKeySpec {
        val spec = PBEKeySpec(password.toCharArray(), salt, PBKDF2_ITERATIONS, KEY_SIZE_BITS)
        val factory = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256")
        val keyBytes = factory.generateSecret(spec).encoded
        return SecretKeySpec(keyBytes, "AES")
    }

    private fun getOrCreateLegacySecretKey(): SecretKey {
        val keyStore = KeyStore.getInstance(ANDROID_KEYSTORE).apply { load(null) }
        (keyStore.getKey(LEGACY_KEY_ALIAS, null) as? SecretKey)?.let { return it }

        val keyGenerator = KeyGenerator.getInstance(KeyProperties.KEY_ALGORITHM_AES, ANDROID_KEYSTORE)
        val spec = KeyGenParameterSpec.Builder(
            LEGACY_KEY_ALIAS,
            KeyProperties.PURPOSE_ENCRYPT or KeyProperties.PURPOSE_DECRYPT
        )
            .setBlockModes(KeyProperties.BLOCK_MODE_GCM)
            .setEncryptionPaddings(KeyProperties.ENCRYPTION_PADDING_NONE)
            .setRandomizedEncryptionRequired(true)
            .setKeySize(256)
            .build()
        keyGenerator.init(spec)
        return keyGenerator.generateKey()
    }
}
