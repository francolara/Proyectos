USE [DbSportCenter]
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Retira de EspaciosDeportivos los Id locales de deporte y suelo despues del cambio de aplicacion y procedimientos.
-- =============================================
-- Firma: FRANCO LARA - 08/10/2026 | Retiro final aprobado en QA; descubre y elimina dependencias fisicas sin asumir nombres de constraints o indices.
DECLARE @RespaldoVerificado BIT = 1;
DECLARE @PruebasQaAprobadas BIT = 1;
DECLARE @ContratosCanonicosPublicados BIT = 1;

IF @RespaldoVerificado <> 1
   OR @PruebasQaAprobadas <> 1
   OR @ContratosCanonicosPublicados <> 1
BEGIN
    RAISERROR('Ejecucion bloqueada: confirme respaldo, pruebas QA y publicacion de contratos canonicos.', 16, 1);
    RETURN;
END;

BEGIN TRY
    IF OBJECT_ID(N'dbo.EspaciosDeportivos', N'U') IS NULL
        RAISERROR('No existe dbo.EspaciosDeportivos.', 16, 1);

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteSuperId') IS NULL
       OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloSuperId') IS NULL
        RAISERROR('No se puede retirar el contrato heredado: faltan las columnas globales.', 16, 1);

    EXEC sys.sp_executesql
        N'
            IF EXISTS
            (
                SELECT 1
                FROM dbo.EspaciosDeportivos e
                WHERE e.TipoDeporteSuperId IS NULL
                   OR e.TipoSueloSuperId IS NULL
            )
                RAISERROR(''No se puede retirar el contrato heredado: hay espacios sin deporte o suelo global.'', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.EspaciosDeportivos e
                LEFT JOIN dbo.TiposDeporteSuperMaestro td ON td.Id = e.TipoDeporteSuperId
                LEFT JOIN dbo.TiposSueloSuperMaestro ts ON ts.Id = e.TipoSueloSuperId
                WHERE td.Id IS NULL
                   OR ts.Id IS NULL
            )
                RAISERROR(''No se puede retirar el contrato heredado: hay identificadores globales sin maestro.'', 16, 1);';

    IF EXISTS
    (
        SELECT 1
        FROM sys.columns c
        WHERE c.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND c.name IN (N'TipoDeporteSuperId', N'TipoSueloSuperId')
          AND c.is_nullable = 1
    )
        RAISERROR('No se puede retirar el contrato heredado: las columnas globales aun permiten NULL; ejecute primero Fase 7A.', 16, 1);

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.foreign_key_columns fkc
        INNER JOIN sys.foreign_keys fk ON fk.object_id = fkc.constraint_object_id
        WHERE fkc.parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoDeporteSuperId'
          AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposDeporteSuperMaestro')
          AND fk.is_disabled = 0
          AND fk.is_not_trusted = 0
    )
       OR NOT EXISTS
    (
        SELECT 1
        FROM sys.foreign_key_columns fkc
        INNER JOIN sys.foreign_keys fk ON fk.object_id = fkc.constraint_object_id
        WHERE fkc.parent_object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
          AND COL_NAME(fkc.parent_object_id, fkc.parent_column_id) = N'TipoSueloSuperId'
          AND fkc.referenced_object_id = OBJECT_ID(N'dbo.TiposSueloSuperMaestro')
          AND fk.is_disabled = 0
          AND fk.is_not_trusted = 0
    )
        RAISERROR('No se puede retirar el contrato heredado: faltan claves foraneas globales confiables.', 16, 1);

    DECLARE @ColumnasHeredadas TABLE
    (
        ObjectId INT NOT NULL,
        ColumnId INT NOT NULL,
        Esquema SYSNAME NOT NULL,
        Tabla SYSNAME NOT NULL,
        Columna SYSNAME NOT NULL,
        PRIMARY KEY (ObjectId, ColumnId)
    );

    INSERT INTO @ColumnasHeredadas (ObjectId, ColumnId, Esquema, Tabla, Columna)
    SELECT
        c.object_id,
        c.column_id,
        s.name,
        t.name,
        c.name
    FROM sys.columns c
    INNER JOIN sys.tables t ON t.object_id = c.object_id
    INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
    WHERE s.name = N'dbo'
      AND t.name = N'EspaciosDeportivos'
      AND c.name IN (N'TipoDeporteId', N'TipoSueloId');

    IF EXISTS
    (
        SELECT 1
        FROM sys.sql_expression_dependencies sed
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = sed.referenced_id
           AND ch.ColumnId = sed.referenced_minor_id
        INNER JOIN sys.objects o ON o.object_id = sed.referencing_id
        WHERE o.type IN (N'P', N'V', N'FN', N'IF', N'TF', N'TR')
    )
       OR EXISTS
    (
        SELECT 1
        FROM sys.sql_modules sm
        INNER JOIN sys.objects o ON o.object_id = sm.object_id
        WHERE o.is_ms_shipped = 0
          AND o.object_id <> OBJECT_ID(N'dbo.Sp_Sistema_ValidarContratoCanonico')
          AND
          (
              sm.definition LIKE N'%.TipoDeporteId%'
              OR sm.definition LIKE N'%.TipoSueloId%'
          )
    )
    BEGIN
        SELECT DISTINCT
            OBJECT_SCHEMA_NAME(o.object_id) AS Esquema,
            o.name AS ObjetoDependiente
        FROM sys.sql_modules sm
        INNER JOIN sys.objects o ON o.object_id = sm.object_id
        WHERE o.is_ms_shipped = 0
          AND o.object_id <> OBJECT_ID(N'dbo.Sp_Sistema_ValidarContratoCanonico')
          AND
          (
              sm.definition LIKE N'%.TipoDeporteId%'
              OR sm.definition LIKE N'%.TipoSueloId%'
              OR EXISTS
              (
                  SELECT 1
                  FROM sys.sql_expression_dependencies sed
                  INNER JOIN @ColumnasHeredadas ch
                      ON ch.ObjectId = sed.referenced_id
                     AND ch.ColumnId = sed.referenced_minor_id
                  WHERE sed.referencing_id = o.object_id
              )
          );

        RAISERROR('No se puede retirar el contrato heredado: aun existen modulos SQL que usan Id locales.', 16, 1);
    END;

    BEGIN TRANSACTION;

    DECLARE @Sql NVARCHAR(MAX);

    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'ALTER TABLE '
                + QUOTENAME(OBJECT_SCHEMA_NAME(fk.parent_object_id)) + N'.'
                + QUOTENAME(OBJECT_NAME(fk.parent_object_id))
                + N' DROP CONSTRAINT ' + QUOTENAME(fk.name) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.foreign_keys fk
        INNER JOIN sys.foreign_key_columns fkc ON fkc.constraint_object_id = fk.object_id
        INNER JOIN @ColumnasHeredadas ch
            ON (ch.ObjectId = fkc.parent_object_id AND ch.ColumnId = fkc.parent_column_id)
            OR (ch.ObjectId = fkc.referenced_object_id AND ch.ColumnId = fkc.referenced_column_id)
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
                + N' DROP CONSTRAINT ' + QUOTENAME(dc.name) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.default_constraints dc
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = dc.parent_object_id
           AND ch.ColumnId = dc.parent_column_id
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
                + N' DROP CONSTRAINT ' + QUOTENAME(cc.name) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.check_constraints cc
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = cc.parent_object_id
        WHERE cc.parent_column_id = ch.ColumnId
           OR EXISTS
           (
               SELECT 1
               FROM sys.sql_expression_dependencies sed
               WHERE sed.referencing_id = cc.object_id
                 AND sed.referenced_id = ch.ObjectId
                 AND sed.referenced_minor_id = ch.ColumnId
           )
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
                + N' DROP CONSTRAINT ' + QUOTENAME(kc.name) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.key_constraints kc
        INNER JOIN sys.index_columns ic
            ON ic.object_id = kc.parent_object_id
           AND ic.index_id = kc.unique_index_id
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = ic.object_id
           AND ch.ColumnId = ic.column_id
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'DROP INDEX ' + QUOTENAME(i.name) + N' ON '
                + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.indexes i
        INNER JOIN sys.index_columns ic
            ON ic.object_id = i.object_id
           AND ic.index_id = i.index_id
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = ic.object_id
           AND ch.ColumnId = ic.column_id
        WHERE i.is_primary_key = 0
          AND i.is_unique_constraint = 0
          AND i.name IS NOT NULL
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(
                N'DROP STATISTICS ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
                + N'.' + QUOTENAME(st.name) + N';'
                AS NVARCHAR(MAX)
            ) AS Comando
        FROM sys.stats st
        INNER JOIN sys.stats_columns sc
            ON sc.object_id = st.object_id
           AND sc.stats_id = st.stats_id
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = sc.object_id
           AND ch.ColumnId = sc.column_id
        WHERE NOT EXISTS
        (
            SELECT 1
            FROM sys.indexes i
            WHERE i.object_id = st.object_id
              AND i.index_id = st.stats_id
        )
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteId') IS NOT NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos DROP COLUMN TipoDeporteId;';

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloId') IS NOT NULL
        EXEC sys.sp_executesql N'ALTER TABLE dbo.EspaciosDeportivos DROP COLUMN TipoSueloId;';

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteId') IS NOT NULL
       OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloId') IS NOT NULL
        RAISERROR('Retiro incompleto: permanecen columnas locales en EspaciosDeportivos.', 16, 1);

    COMMIT TRANSACTION;

    SELECT
        N'FASE_7B_COMPLETADA' AS Resultado,
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
