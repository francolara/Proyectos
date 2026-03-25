package com.prestamos.app.data.local

import android.content.Context
import androidx.room.Database
import androidx.room.Room
import androidx.room.RoomDatabase
import androidx.room.TypeConverters
import androidx.room.migration.Migration
import androidx.sqlite.db.SupportSQLiteDatabase
import com.prestamos.app.data.local.dao.ClienteDao
import com.prestamos.app.data.local.dao.CuotaDao
import com.prestamos.app.data.local.dao.PagoDao
import com.prestamos.app.data.local.dao.PrestamoDao
import com.prestamos.app.data.local.dao.TipoCobroDao
import com.prestamos.app.data.local.entity.ClienteEntity
import com.prestamos.app.data.local.entity.CuotaEntity
import com.prestamos.app.data.local.entity.PagoEntity
import com.prestamos.app.data.local.entity.PrestamoEntity
import com.prestamos.app.data.local.entity.TipoCobroEntity

@Database(
    entities = [
        ClienteEntity::class,
        PrestamoEntity::class,
        CuotaEntity::class,
        PagoEntity::class,
        TipoCobroEntity::class
    ],
    version = 4,
    exportSchema = false
)
@TypeConverters(Converters::class)
abstract class AppDatabase : RoomDatabase() {
    abstract fun clienteDao(): ClienteDao
    abstract fun prestamoDao(): PrestamoDao
    abstract fun cuotaDao(): CuotaDao
    abstract fun pagoDao(): PagoDao
    abstract fun tipoCobroDao(): TipoCobroDao

    companion object {
        const val DATABASE_NAME = "prestamos.db"
        const val DATABASE_VERSION = 4

        val MIGRATION_1_2 = object : Migration(1, 2) {
            override fun migrate(db: SupportSQLiteDatabase) {
                addColumnIfMissing(db, "clientes", "fechaModificacion", "INTEGER NOT NULL DEFAULT 0")
                addColumnIfMissing(db, "prestamos", "fechaModificacion", "INTEGER NOT NULL DEFAULT 0")
                addColumnIfMissing(db, "cuotas", "fechaModificacion", "INTEGER NOT NULL DEFAULT 0")
                addColumnIfMissing(db, "pagos", "fechaModificacion", "INTEGER NOT NULL DEFAULT 0")
                addColumnIfMissing(db, "pagos", "observacion", "TEXT")

                db.execSQL("CREATE UNIQUE INDEX IF NOT EXISTS index_clientes_documentoIdentidad ON clientes(documentoIdentidad)")
                db.execSQL("CREATE INDEX IF NOT EXISTS index_prestamos_idCliente ON prestamos(idCliente)")
                db.execSQL("CREATE INDEX IF NOT EXISTS index_cuotas_idPrestamo ON cuotas(idPrestamo)")
                db.execSQL("CREATE INDEX IF NOT EXISTS index_pagos_idPrestamo ON pagos(idPrestamo)")
                db.execSQL("CREATE INDEX IF NOT EXISTS index_pagos_idCuota ON pagos(idCuota)")
            }
        }

        val MIGRATION_2_3 = object : Migration(2, 3) {
            override fun migrate(db: SupportSQLiteDatabase) {
                db.execSQL(
                    """
                    CREATE TABLE IF NOT EXISTS tipos_cobro (
                        idTipoCobro INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
                        nombre TEXT NOT NULL,
                        fechaRegistro INTEGER NOT NULL,
                        fechaModificacion INTEGER NOT NULL
                    )
                    """.trimIndent()
                )
                db.execSQL("CREATE UNIQUE INDEX IF NOT EXISTS index_tipos_cobro_nombre ON tipos_cobro(nombre)")
                addColumnIfMissing(db, "pagos", "idTipoCobro", "INTEGER")
                db.execSQL("CREATE INDEX IF NOT EXISTS index_pagos_idTipoCobro ON pagos(idTipoCobro)")
            }
        }

        val MIGRATION_3_4 = object : Migration(3, 4) {
            override fun migrate(db: SupportSQLiteDatabase) {
                addColumnIfMissing(db, "cuotas", "moraPendiente", "REAL NOT NULL DEFAULT 0")
                addColumnIfMissing(db, "pagos", "moraCobrada", "REAL NOT NULL DEFAULT 0")
            }
        }

        @Volatile
        private var INSTANCE: AppDatabase? = null

        fun getInstance(context: Context): AppDatabase {
            return INSTANCE ?: synchronized(this) {
                val instance = Room.databaseBuilder(
                    context.applicationContext,
                    AppDatabase::class.java,
                    DATABASE_NAME
                )
                    .addMigrations(MIGRATION_1_2, MIGRATION_2_3, MIGRATION_3_4)
                    .fallbackToDestructiveMigrationOnDowngrade(false)
                    .build()
                INSTANCE = instance
                instance
            }
        }

        fun closeInstance() {
            synchronized(this) {
                INSTANCE?.close()
                INSTANCE = null
            }
        }

        private fun addColumnIfMissing(
            db: SupportSQLiteDatabase,
            tableName: String,
            columnName: String,
            columnDefinition: String
        ) {
            db.query("PRAGMA table_info($tableName)").use { cursor ->
                var found = false
                val nameIndex = cursor.getColumnIndex("name")
                while (cursor.moveToNext()) {
                    if (nameIndex >= 0 && cursor.getString(nameIndex).equals(columnName, ignoreCase = true)) {
                        found = true
                        break
                    }
                }
                if (!found) {
                    db.execSQL("ALTER TABLE $tableName ADD COLUMN $columnName $columnDefinition")
                }
            }
        }
    }
}
