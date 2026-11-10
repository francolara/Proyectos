-- Firma: FRANCO LARA | Create date: 09/11/2026 | Catalogos peruanos y SCTR por trabajador.
USE [SistemaWebHYT];
GO
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO

-- Fuente: Tablas/BancoPeruHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   BancoPeruHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.BancoPeruHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.BancoPeruHYT
    (
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT PK_BancoPeruHYT PRIMARY KEY
    );
END;
GO

-- Fuente: Tablas/EntidadPrevisionalHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   EntidadPrevisionalHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.EntidadPrevisionalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EntidadPrevisionalHYT
    (
        Nombre NVARCHAR(100) NOT NULL CONSTRAINT PK_EntidadPrevisionalHYT PRIMARY KEY,
        Regimen VARCHAR(3) NOT NULL CONSTRAINT CK_EntidadPrevisionalHYT CHECK (Regimen IN ('AFP', 'ONP'))
    );
END;
GO

-- Fuente: Tablas/PolizaSctrTrabajadorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PolizaSctrTrabajadorHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PolizaSctrTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PolizaSctrTrabajadorHYT
    (
        PolizaId BIGINT NOT NULL CONSTRAINT FK_PolizaSctrTrabajadorHYT_Poliza REFERENCES dbo.PolizaSctrHYT(PolizaId),
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PolizaSctrTrabajadorHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        CONSTRAINT PK_PolizaSctrTrabajadorHYT PRIMARY KEY (PolizaId, TrabajadorId)
    );
END;
GO

-- Fuente: Script/041_CatalogosBancosPrevisionHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Carga entidades bancarias y previsionales para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
-- Fuentes: directorio de captación de depósitos y listado de AFP de la SBS.
INSERT INTO dbo.BancoPeruHYT (Nombre)
SELECT nuevo.Nombre FROM (VALUES
    (N'Alfin Banco'),
    (N'BanBif'),
    (N'Banco Bci'),
    (N'Banco de Crédito del Perú (BCP)'),
    (N'Banco Efectiva'),
    (N'Banco Falabella'),
    (N'Banco GNB'),
    (N'Banco Pichincha'),
    (N'Banco Ripley'),
    (N'Bancom'),
    (N'Bank of China (Perú)'),
    (N'BBVA Perú'),
    (N'Santander Consumer'),
    (N'Citibank del Perú'),
    (N'Compartamos Banco'),
    (N'ICBC Peru Bank'),
    (N'Interbank'),
    (N'Mibanco'),
    (N'Banco de la Nación'),
    (N'Santander Perú'),
    (N'Scotiabank Perú')
) AS nuevo (Nombre)
WHERE NOT EXISTS (SELECT 1 FROM dbo.BancoPeruHYT AS banco WHERE banco.Nombre = nuevo.Nombre);

INSERT INTO dbo.EntidadPrevisionalHYT (Nombre, Regimen)
SELECT nuevo.Nombre, nuevo.Regimen FROM (VALUES
    (N'AFP Hábitat', 'AFP'), (N'AFP Integra', 'AFP'), (N'Prima AFP', 'AFP'),
    (N'Profuturo AFP', 'AFP'), (N'ONP', 'ONP')
) AS nuevo (Nombre, Regimen)
WHERE NOT EXISTS (SELECT 1 FROM dbo.EntidadPrevisionalHYT AS entidad WHERE entidad.Nombre = nuevo.Nombre);
GO

-- Fuente: StoreProcedure/spu_CatalogosPlanillasHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Devuelve bancos y entidades previsionales de referencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CatalogosPlanillasHYT_Consultar
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51000, N'No tiene permiso para consultar catalogos.', 1;
        SELECT (SELECT JSON_QUERY((SELECT banco.Nombre FROM dbo.BancoPeruHYT AS banco ORDER BY banco.Nombre FOR JSON PATH)) AS Bancos,
            JSON_QUERY((SELECT entidad.Nombre, entidad.Regimen FROM dbo.EntidadPrevisionalHYT AS entidad
                ORDER BY entidad.Regimen, entidad.Nombre FOR JSON PATH)) AS Entidades
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/042_PermisosCatalogosPlanillasHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Habilita consulta de catálogos de referencia.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CatalogosPlanillasHYT_Consultar TO hyt_app_executor;
GO

-- Fuente: StoreProcedure/spu_ObligacionesHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta obligaciones y polizas con sus trabajadores cubiertos.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObligacionesHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar obligaciones.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        SELECT (
            SELECT @Mes AS Mes,
                   JSON_QUERY((
                       SELECT obligacion.ObligacionId, obligacion.Tipo, obligacion.Entidad, obligacion.Mes,
                              obligacion.Importe, obligacion.Vencimiento, obligacion.FechaPago,
                              obligacion.ReferenciaPago, obligacion.VersionFila,
                              CASE WHEN obligacion.FechaPago IS NOT NULL THEN 'PAGADA'
                                   WHEN obligacion.Vencimiento < @Hoy THEN 'VENCIDA' ELSE 'PENDIENTE' END AS Estado
                       FROM dbo.ObligacionPersonalHYT AS obligacion
                       WHERE obligacion.Mes = @Mes
                       ORDER BY obligacion.Vencimiento, obligacion.Tipo, obligacion.Entidad
                       FOR JSON PATH
                   )) AS Obligaciones,
                   JSON_QUERY((
                       SELECT poliza.PolizaId, poliza.Aseguradora, poliza.NumeroPoliza, poliza.Cobertura,
                              poliza.VigenciaDesde, poliza.VigenciaHasta, poliza.VersionFila,
                              JSON_QUERY(COALESCE((
                                  SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre
                                  FROM dbo.PolizaSctrTrabajadorHYT AS cobertura
                                  INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = cobertura.TrabajadorId
                                  WHERE cobertura.PolizaId = poliza.PolizaId
                                  ORDER BY trabajador.NombreCompleto FOR JSON PATH
                              ), N'[]')) AS Trabajadores,
                              CASE WHEN poliza.VigenciaHasta < @Hoy THEN 'VENCIDA'
                                   WHEN poliza.VigenciaDesde > @Hoy THEN 'PROGRAMADA'
                                   WHEN DATEDIFF(DAY, @Hoy, poliza.VigenciaHasta) <= 30 THEN 'POR_VENCER'
                                   ELSE 'VIGENTE' END AS Estado
                       FROM dbo.PolizaSctrHYT AS poliza
                       ORDER BY poliza.VigenciaHasta DESC, poliza.PolizaId DESC
                       FOR JSON PATH
                   )) AS Polizas,
                   JSON_QUERY(COALESCE((
                       SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre
                       FROM dbo.TrabajadorHYT AS trabajador WHERE trabajador.Activo = 1
                       ORDER BY trabajador.NombreCompleto FOR JSON PATH
                   ), N'[]')) AS PersonalActivo
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PolizaSctrHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda poliza SCTR y trabajadores cubiertos en una sola transaccion.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PolizaSctrHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Aseguradora NVARCHAR(120),
    @NumeroPoliza NVARCHAR(80),
    @Cobertura VARCHAR(10),
    @VigenciaDesde DATE,
    @VigenciaHasta DATE,
    @PolizaId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @Trabajadores NVARCHAR(MAX) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar polizas SCTR.', 1;

        SET @Aseguradora = NULLIF(TRIM(@Aseguradora), N'');
        SET @NumeroPoliza = NULLIF(TRIM(@NumeroPoliza), N'');
        IF @Aseguradora IS NULL OR @NumeroPoliza IS NULL OR @Cobertura IS NULL
           OR @Cobertura NOT IN ('SALUD', 'PENSION') OR @VigenciaDesde IS NULL
           OR @VigenciaHasta IS NULL OR @VigenciaHasta < @VigenciaDesde
            THROW 51000, N'Revise la aseguradora, la cobertura y la vigencia de la poliza.', 1;

        DECLARE @Personal TABLE (TrabajadorId BIGINT PRIMARY KEY);
        IF @Trabajadores IS NOT NULL
        BEGIN
            IF COALESCE(ISJSON(@Trabajadores, ARRAY), 0) <> 1
                THROW 51000, N'Seleccione trabajadores validos.', 1;

            INSERT INTO @Personal (TrabajadorId)
            SELECT DISTINCT TRY_CONVERT(BIGINT, fila.value)
            FROM OPENJSON(@Trabajadores) AS fila
            WHERE TRY_CONVERT(BIGINT, fila.value) > 0;
            IF NOT EXISTS (SELECT 1 FROM @Personal)
                THROW 51000, N'Seleccione al menos un trabajador.', 1;
        END;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @PolizaId IS NULL
        BEGIN
            INSERT INTO dbo.PolizaSctrHYT (Aseguradora, NumeroPoliza, Cobertura, VigenciaDesde, VigenciaHasta)
            VALUES (@Aseguradora, @NumeroPoliza, @Cobertura, @VigenciaDesde, @VigenciaHasta);
            SET @PolizaId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.PolizaSctrHYT
            SET Aseguradora = @Aseguradora, NumeroPoliza = @NumeroPoliza, Cobertura = @Cobertura,
                VigenciaDesde = @VigenciaDesde, VigenciaHasta = @VigenciaHasta
            WHERE PolizaId = @PolizaId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La poliza cambio o no existe. Vuelva a abrirla.', 1;
        END;

        IF @Trabajadores IS NOT NULL
        BEGIN
            IF EXISTS (
                SELECT 1 FROM @Personal AS seleccionado
                LEFT JOIN dbo.TrabajadorHYT AS trabajador WITH (HOLDLOCK) ON trabajador.TrabajadorId = seleccionado.TrabajadorId
                LEFT JOIN dbo.PolizaSctrTrabajadorHYT AS existente
                    ON existente.PolizaId = @PolizaId AND existente.TrabajadorId = seleccionado.TrabajadorId
                WHERE trabajador.TrabajadorId IS NULL OR (trabajador.Activo = 0 AND existente.TrabajadorId IS NULL))
                THROW 51000, N'Uno de los trabajadores seleccionados no esta disponible.', 1;

            DELETE FROM dbo.PolizaSctrTrabajadorHYT WHERE PolizaId = @PolizaId;
            INSERT INTO dbo.PolizaSctrTrabajadorHYT (PolizaId, TrabajadorId)
            SELECT @PolizaId, seleccionado.TrabajadorId FROM @Personal AS seleccionado;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PolizaId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Aseguradora AS Aseguradora, @NumeroPoliza AS NumeroPoliza, @Cobertura AS Cobertura,
                   @VigenciaDesde AS VigenciaDesde, @VigenciaHasta AS VigenciaHasta, JSON_QUERY(@Trabajadores) AS Trabajadores
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PolizaSctrHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @PolizaId AS PolizaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

COMMIT TRANSACTION;
GO
