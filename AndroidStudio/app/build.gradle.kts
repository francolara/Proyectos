import java.util.Properties
import java.io.File

plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.compose)
    id("com.google.devtools.ksp")
}

val versionPropertiesFile = rootProject.file("version.properties")
val versionProperties = Properties().apply {
    if (versionPropertiesFile.exists()) {
        versionPropertiesFile.inputStream().use(::load)
    } else {
        setProperty("VERSION_CODE", "1")
        setProperty("VERSION_NAME", "1.0.0")
        versionPropertiesFile.outputStream().use { store(it, "Configuracion de version de la app") }
    }
}

val isReleaseBuildRequested = gradle.startParameter.taskNames.any { taskName ->
    val lower = taskName.lowercase()
    lower.contains("release") && (
        lower.contains("assemble") ||
            lower.contains("bundle") ||
            lower.contains("package") ||
            lower.contains("install")
        )
}
val isDebugBuildRequested = gradle.startParameter.taskNames.any { taskName ->
    val lower = taskName.lowercase()
    lower.contains("debug") && (
        lower.contains("assemble") ||
            lower.contains("install") ||
            lower.contains("package")
        )
}

var computedVersionCode = versionProperties.getProperty("VERSION_CODE", "1").toIntOrNull() ?: 1
var computedVersionName = versionProperties.getProperty("VERSION_NAME", "1.0.0")

fun bumpPatchVersion(versionName: String): String {
    val parts = versionName.split(".")
    if (parts.size < 3) return "1.0.1"
    val major = parts[0].toIntOrNull() ?: 1
    val minor = parts[1].toIntOrNull() ?: 0
    val patch = parts[2].toIntOrNull() ?: 0
    return "$major.$minor.${patch + 1}"
}

if (isReleaseBuildRequested) {
    computedVersionCode += 1
    computedVersionName = bumpPatchVersion(computedVersionName)
    versionProperties.setProperty("VERSION_CODE", computedVersionCode.toString())
    versionProperties.setProperty("VERSION_NAME", computedVersionName)
    versionPropertiesFile.outputStream().use { versionProperties.store(it, "Configuracion de version de la app") }
} else if (isDebugBuildRequested) {
    val debugAutoCode = (System.currentTimeMillis() / 60000L).toInt()
    computedVersionCode = maxOf(computedVersionCode, debugAutoCode)
    computedVersionName = "$computedVersionName-debug.$computedVersionCode"
}

android {
    namespace = "com.prestamos.app"
    compileSdk {
        version = release(36) {
            minorApiLevel = 1
        }
    }

    defaultConfig {
        applicationId = "com.prestamos.app"
        minSdk = 24
        targetSdk = 36
        // Versionado:
        // - Release: VERSION_CODE sube automaticamente en assembleRelease/bundleRelease.
        // - Debug: VERSION_CODE usa marca de tiempo (minutos) para actualizar siempre.
        // - VERSION_NAME base se define en AndroidStudio/version.properties.
        versionCode = computedVersionCode
        versionName = computedVersionName

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_11
        targetCompatibility = JavaVersion.VERSION_11
        isCoreLibraryDesugaringEnabled = true
    }
    buildFeatures {
        compose = true
        buildConfig = true
    }
}

fun copyVersionedApk(buildType: String) {
    val safeVersionName = computedVersionName.replace(Regex("[^A-Za-z0-9._-]"), "_")
    val apkDir = layout.buildDirectory.dir("outputs/apk/$buildType").get().asFile
    if (!apkDir.exists()) return
    val targetName = "AppPrestamos-$buildType-v${safeVersionName}-vc${computedVersionCode}.apk"
    apkDir.listFiles()
        ?.filter { it.isFile && it.extension.equals("apk", ignoreCase = true) }
        ?.forEach { apk ->
            if (apk.name != targetName) {
                apk.copyTo(File(apkDir, targetName), overwrite = true)
            }
        }
}

tasks.matching { it.name == "assembleRelease" }.configureEach {
    doLast { copyVersionedApk("release") }
}

tasks.matching { it.name == "assembleDebug" }.configureEach {
    doLast { copyVersionedApk("debug") }
}

dependencies {
    coreLibraryDesugaring("com.android.tools:desugar_jdk_libs:2.1.5")
    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.lifecycle.runtime.ktx)
    implementation("androidx.lifecycle:lifecycle-runtime-compose:2.8.6")
    implementation(libs.androidx.activity.compose)
    implementation(platform("androidx.compose:compose-bom:2024.09.00"))
    implementation(libs.androidx.compose.ui)
    implementation(libs.androidx.compose.ui.graphics)
    implementation(libs.androidx.compose.ui.tooling.preview)
    implementation(libs.androidx.compose.material3)
    implementation("androidx.compose.material:material-icons-extended")
    implementation("androidx.navigation:navigation-compose:2.8.2")
    implementation("androidx.lifecycle:lifecycle-viewmodel-compose:2.8.6")
    implementation("androidx.datastore:datastore-preferences:1.1.1")
    implementation("androidx.documentfile:documentfile:1.1.0")
    implementation("androidx.biometric:biometric:1.1.0")
    implementation("androidx.work:work-runtime-ktx:2.10.3")
    testImplementation(libs.junit)
    androidTestImplementation(libs.androidx.junit)
    androidTestImplementation(libs.androidx.espresso.core)
    androidTestImplementation(platform("androidx.compose:compose-bom:2024.09.00"))
    androidTestImplementation(libs.androidx.compose.ui.test.junit4)
    debugImplementation(libs.androidx.compose.ui.tooling)
    debugImplementation(libs.androidx.compose.ui.test.manifest)

    // ROOM DATABASE
    implementation("androidx.room:room-runtime:2.8.4")
    implementation("androidx.room:room-ktx:2.8.4")
    ksp("androidx.room:room-compiler:2.8.4")
}
