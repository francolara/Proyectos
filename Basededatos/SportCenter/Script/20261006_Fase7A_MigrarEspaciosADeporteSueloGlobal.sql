USE [DbSportCenter]
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Agrega, completa y protege los identificadores globales de deporte y suelo en EspaciosDeportivos.
-- =============================================
-- Firma: FRANCO LARA - 08/10/2026 | Corrige la migracion para bases sin columnas globales, valida equivalencia y deja los Id locales anulables durante el despliegue coordinado.
BEGIN TRY
    BEGIN TRANSACTION;

    IF OBJECT_ID(N'dbo.EspaciosDeportivos', N'U') IS NULL
        RAISERROR('No existe dbo.EspaciosDeportivos.', 16, 1);

    IF OBJECT_ID(N'dbo.TiposDeporteSuperMaestro', N'U') IS NULL
       OR OBJECT_ID(N'dbo.TiposSueloSuperMaestro', N'U') IS NULL
       OR OBJECT_ID(N'dbo.TiposDeporte', N'U') IS NULL
       OR OBJECT_ID(N'dbo.TiposSuelo', N'U') IS NULL
        RAISERROR('Faltan tablas de maestros requeridas para la migracion.', 16, 1);

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteSuperId') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos ADD TipoDeporteSuperId INT NULL;';

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloSuperId') IS NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos ADD TipoSueloSuperId INT NULL;';

    DECLARE @DeportesMigrados INT = 0;
    DECLARE @SuelosMigrados INT = 0;

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteId') IS NOT NULL
    BEGIN
        EXEC sys.sp_executesql
            N'
                UPDATE e
                SET e.TipoDeporteSuperId = td.TipoDeporteSuperId
                FROM dbo.EspaciosDeportivos e
                INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE e.TipoDeporteSuperId IS NULL
                  AND td.NegocioId = s.NegocioId
                  AND td.TipoDeporteSuperId IS NOT NULL;

                SET @FilasMigradas = @@ROWCOUNT;

                IF EXISTS
                (
                    SELECT 1
                    FROM dbo.EspaciosDeportivos e
                    INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
                    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                    WHERE td.NegocioId <> s.NegocioId
                       OR td.TipoDeporteSuperId IS NULL
                       OR e.TipoDeporteSuperId <> td.TipoDeporteSuperId
                )
                    RAISERROR(''Migracion detenida: TipoDeporteId no coincide con TipoDeporteSuperId o con el negocio del espacio.'', 16, 1);',
            N'@FilasMigradas INT OUTPUT',
            @FilasMigradas = @DeportesMigrados OUTPUT;
    END;

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloId') IS NOT NULL
    BEGIN
        EXEC sys.sp_executesql
            N'
                UPDATE e
                SET e.TipoSueloSuperId = ts.TipoSueloSuperId
                FROM dbo.EspaciosDeportivos e
                INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE e.TipoSueloSuperId IS NULL
                  AND ts.NegocioId = s.NegocioId
                  AND ts.TipoSueloSuperId IS NOT NULL;

                SET @FilasMigradas = @@ROWCOUNT;

                IF EXISTS
                (
                    SELECT 1
                    FROM dbo.EspaciosDeportivos e
                    INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
                    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                    WHERE ts.NegocioId <> s.NegocioId
                       OR ts.TipoSueloSuperId IS NULL
                       OR e.TipoSueloSuperId <> ts.TipoSueloSuperId
                )
                    RAISERROR(''Migracion detenida: TipoSueloId no coincide con TipoSueloSuperId o con el negocio del espacio.'', 16, 1);',
            N'@FilasMigradas INT OUTPUT',
            @FilasMigradas = @SuelosMigrados OUTPUT;
    END;

    EXEC sys.sp_executesql
        N'
            IF EXISTS
            (
                SELECT 1
                FROM dbo.EspaciosDeportivos e
                WHERE e.TipoDeporteSuperId IS NULL
                   OR e.TipoSueloSuperId IS NULL
            )
                RAISERROR(''No se puede completar la migracion: existen espacios sin mapeo global de deporte o suelo.'', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.EspaciosDeportivos e
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                LEFT JOIN dbo.TiposDeporteSuperMaestro tdm ON tdm.Id = e.TipoDeporteSuperId
                LEFT JOIN dbo.TiposSueloSuperMaestro tsm ON tsm.Id = e.TipoSueloSuperId
                WHERE tdm.Id IS NULL
                   OR tsm.Id IS NULL
                   OR NOT EXISTS
                   (
                       SELECT 1
                       FROM dbo.TiposDeporte td
                       WHERE td.NegocioId = s.NegocioId
                         AND td.TipoDeporteSuperId = e.TipoDeporteSuperId
                         AND td.Activo = 1
                   )
                   OR NOT EXISTS
                   (
                       SELECT 1
                       FROM dbo.TiposSuelo ts
                       WHERE ts.NegocioId = s.NegocioId
                         AND ts.TipoSueloSuperId = e.TipoSueloSuperId
                         AND ts.Activo = 1
                   )
            )
                RAISERROR(''No se puede completar la migracion: el deporte o suelo global no pertenece a la configuracion activa del negocio.'', 16, 1);

            ALTER TABLE dbo.EspaciosDeportivos ALTER COLUMN TipoDeporteSuperId INT NOT NULL;
            ALTER TABLE dbo.EspaciosDeportivos ALTER COLUMN TipoSueloSuperId INT NOT NULL;';

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteId') IS NOT NULL
       AND EXISTS
       (
           SELECT 1
           FROM sys.columns c
           WHERE c.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
             AND c.name = N'TipoDeporteId'
             AND c.is_nullable = 0
       )
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos ALTER COLUMN TipoDeporteId INT NULL;';

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloId') IS NOT NULL
       AND EXISTS
       (
           SELECT 1
           FROM sys.columns c
           WHERE c.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
             AND c.name = N'TipoSueloId'
             AND c.is_nullable = 0
       )
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos ALTER COLUMN TipoSueloId INT NULL;';

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.foreign_key_columns fkc
        WHERE fkc.parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoDeporteSuperId'
          AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposDeporteSuperMaestro')
          AND COL_NAME(fkc.referenced_object_id, fkc.referenced_column_id) = N'Id'
    )
    BEGIN
        EXEC sys.sp_executesql
            N'ALTER TABLE dbo.EspaciosDeportivos WITH CHECK
              ADD CONSTRAINT FK_EspaciosDeportivos_TiposDeporteSuperMaestro_TipoDeporteSuperId
                  FOREIGN KEY (TipoDeporteSuperId) REFERENCES dbo.TiposDeporteSuperMaestro(Id);
              ALTER TABLE dbo.EspaciosDeportivos
              CHECK CONSTRAINT FK_EspaciosDeportivos_TiposDeporteSuperMaestro_TipoDeporteSuperId;';
    END;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.foreign_key_columns fkc
        WHERE fkc.parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoSueloSuperId'
          AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposSueloSuperMaestro')
          AND COL_NAME(fkc.referenced_object_id, fkc.referenced_column_id) = N'Id'
    )
    BEGIN
        EXEC sys.sp_executesql
            N'ALTER TABLE dbo.EspaciosDeportivos WITH CHECK
              ADD CONSTRAINT FK_EspaciosDeportivos_TiposSueloSuperMaestro_TipoSueloSuperId
                  FOREIGN KEY (TipoSueloSuperId) REFERENCES dbo.TiposSueloSuperMaestro(Id);
              ALTER TABLE dbo.EspaciosDeportivos
              CHECK CONSTRAINT FK_EspaciosDeportivos_TiposSueloSuperMaestro_TipoSueloSuperId;';
    END;

    DECLARE @SqlConfiarClaves NVARCHAR(MAX);

    SELECT @SqlConfiarClaves = STRING_AGG(
        CAST(
            N'ALTER TABLE dbo.EspaciosDeportivos WITH CHECK CHECK CONSTRAINT '
            + QUOTENAME(fk.name) + N';'
            AS NVARCHAR(MAX)
        ),
        NCHAR(10)
    )
    FROM sys.foreign_keys fk
    INNER JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
    WHERE fkc.parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
      AND
      (
          (
              COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoDeporteSuperId'
              AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposDeporteSuperMaestro')
          )
          OR
          (
              COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoSueloSuperId'
              AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposSueloSuperMaestro')
          )
      );

    IF NULLIF(@SqlConfiarClaves, N'') IS NOT NULL
        EXEC sys.sp_executesql @SqlConfiarClaves;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.indexes i
        INNER JOIN sys.index_columns ic
            ON ic.object_id = i.object_id
           AND ic.index_id = i.index_id
        WHERE i.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND ic.key_ordinal = 1
          AND COL_NAME(ic.object_id, ic.column_id) = N'TipoDeporteSuperId'
    )
        EXEC sys.sp_executesql N'CREATE INDEX IX_EspaciosDeportivos_TipoDeporteSuperId ON dbo.EspaciosDeportivos(TipoDeporteSuperId);';

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.indexes i
        INNER JOIN sys.index_columns ic
            ON ic.object_id = i.object_id
           AND ic.index_id = i.index_id
        WHERE i.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND ic.key_ordinal = 1
          AND COL_NAME(ic.object_id, ic.column_id) = N'TipoSueloSuperId'
    )
        EXEC sys.sp_executesql N'CREATE INDEX IX_EspaciosDeportivos_TipoSueloSuperId ON dbo.EspaciosDeportivos(TipoSueloSuperId);';

    COMMIT TRANSACTION;

    SELECT
        N'FASE_7A_COMPLETADA' AS Resultado,
        @DeportesMigrados AS DeportesMigrados,
        @SuelosMigrados AS SuelosMigrados,
        SYSUTCDATETIME() AS FechaUtc;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
GO
