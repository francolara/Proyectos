-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Elimina fisicamente IDs locales despues de publicar y validar contratos canonicos.
-- =============================================
-- Requisitos:
-- 1. Backup verificado.
-- 2. Script 4B_01 ejecutado.
-- 3. SP de Fase 4B y aplicacion canonica publicados y validados.

SET NOCOUNT ON;
SET XACT_ABORT ON;

DECLARE @RespaldoVerificado BIT = 1;
DECLARE @ContratosCanonicosPublicados BIT = 1;

IF @RespaldoVerificado <> 1 OR @ContratosCanonicosPublicados <> 1
BEGIN
    RAISERROR('Ejecucion bloqueada: confirme respaldo y publicacion de contratos canonicos.', 16, 1);
    RETURN;
END;

BEGIN TRY
    IF EXISTS
    (
        SELECT 1
        FROM dbo.Negocios n
        WHERE n.CodigoMoneda IS NULL
           OR LTRIM(RTRIM(n.CodigoMoneda)) = N''
    )
        RAISERROR('Fase 4B bloqueada: existen negocios sin CodigoMoneda.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.ComprobantesElectronicos ce
        WHERE ce.CodigoMoneda IS NULL
           OR LTRIM(RTRIM(ce.CodigoMoneda)) = N''
           OR ce.CodigoTipoComprobante IS NULL
           OR LTRIM(RTRIM(ce.CodigoTipoComprobante)) = N''
    )
        RAISERROR('Fase 4B bloqueada: existen comprobantes sin codigos canonicos.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM sys.columns c
        WHERE c.object_id = OBJECT_ID(N'dbo.Negocios')
          AND c.name = N'CodigoMoneda'
          AND c.is_nullable = 1
    )
        RAISERROR('Fase 4B bloqueada: Negocios.CodigoMoneda aun permite valores nulos; ejecute primero 4B_01.', 16, 1);

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.indexes i
        WHERE i.object_id = OBJECT_ID(N'dbo.ComprobantesElectronicos')
          AND i.name = N'UX_ComprobantesElectronicos_Negocio_CodigoTipo_Serie_Numero'
          AND i.is_unique = 1
          AND i.is_disabled = 0
    )
        RAISERROR('Fase 4B bloqueada: falta el indice unico canonico; ejecute primero 4B_01.', 16, 1);

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
        o.name,
        c.name
    FROM sys.columns c
    INNER JOIN sys.objects o ON o.object_id = c.object_id AND o.type = N'U'
    INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
    WHERE (s.name = N'dbo' AND o.name = N'Negocios' AND c.name = N'MonedaId')
       OR (s.name = N'dbo' AND o.name = N'ComprobantesElectronicos' AND c.name IN (N'TipoComprobante', N'TipoMoneda'));

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
    BEGIN
        SELECT DISTINCT
            OBJECT_SCHEMA_NAME(sed.referencing_id) AS Esquema,
            OBJECT_NAME(sed.referencing_id) AS ObjetoDependiente,
            ch.Tabla,
            ch.Columna
        FROM sys.sql_expression_dependencies sed
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = sed.referenced_id
           AND ch.ColumnId = sed.referenced_minor_id
        INNER JOIN sys.objects o ON o.object_id = sed.referencing_id
        WHERE o.type IN (N'P', N'V', N'FN', N'IF', N'TF', N'TR');

        RAISERROR('Fase 4B bloqueada: aun existen modulos SQL dependientes de columnas heredadas.', 16, 1);
    END;

    BEGIN TRANSACTION;

    DECLARE @Sql NVARCHAR(MAX);

    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(N'ALTER TABLE ' + QUOTENAME(OBJECT_SCHEMA_NAME(fk.parent_object_id)) + N'.' + QUOTENAME(OBJECT_NAME(fk.parent_object_id))
            + N' DROP CONSTRAINT ' + QUOTENAME(fk.name) + N';' AS NVARCHAR(MAX)) AS Comando
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
            CAST(N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
            + N' DROP CONSTRAINT ' + QUOTENAME(dc.name) + N';' AS NVARCHAR(MAX)) AS Comando
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
            CAST(N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
            + N' DROP CONSTRAINT ' + QUOTENAME(cc.name) + N';' AS NVARCHAR(MAX)) AS Comando
        FROM sys.check_constraints cc
        INNER JOIN sys.sql_expression_dependencies sed ON sed.referencing_id = cc.object_id
        INNER JOIN @ColumnasHeredadas ch
            ON ch.ObjectId = sed.referenced_id
           AND ch.ColumnId = sed.referenced_minor_id
    ) cmd;
    IF NULLIF(@Sql, N'') IS NOT NULL EXEC sys.sp_executesql @Sql;

    SET @Sql = NULL;
    SELECT @Sql = STRING_AGG(cmd.Comando, NCHAR(10))
    FROM
    (
        SELECT DISTINCT
            CAST(N'ALTER TABLE ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
            + N' DROP CONSTRAINT ' + QUOTENAME(kc.name) + N';' AS NVARCHAR(MAX)) AS Comando
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
            CAST(N'DROP INDEX ' + QUOTENAME(i.name) + N' ON '
            + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla) + N';' AS NVARCHAR(MAX)) AS Comando
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
            CAST(N'DROP STATISTICS ' + QUOTENAME(ch.Esquema) + N'.' + QUOTENAME(ch.Tabla)
            + N'.' + QUOTENAME(st.name) + N';' AS NVARCHAR(MAX)) AS Comando
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

    IF COL_LENGTH(N'dbo.Negocios', N'MonedaId') IS NOT NULL
        ALTER TABLE dbo.Negocios DROP COLUMN MonedaId;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoComprobante') IS NOT NULL
        ALTER TABLE dbo.ComprobantesElectronicos DROP COLUMN TipoComprobante;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoMoneda') IS NOT NULL
        ALTER TABLE dbo.ComprobantesElectronicos DROP COLUMN TipoMoneda;

    COMMIT TRANSACTION;

    SELECT
        N'FASE_4B_COMPLETADA' AS Resultado,
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
