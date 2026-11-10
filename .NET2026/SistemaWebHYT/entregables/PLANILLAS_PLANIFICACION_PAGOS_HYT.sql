-- Firma: FRANCO LARA | Create date: 09/11/2026 | Configuracion del trabajador y planificacion de pagos.
USE [SistemaWebHYT];
GO
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO

-- Fuente: Script/039_ConfiguracionPagoTrabajador.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega horario habitual, regimen previsional y frecuencia de pago.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'HoraEntradaHabitual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD HoraEntradaHabitual TIME(0) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'HoraSalidaHabitual') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD HoraSalidaHabitual TIME(0) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'RegimenPrevisional') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD RegimenPrevisional VARCHAR(3) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'AdministradoraAFP') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD AdministradoraAFP NVARCHAR(100) NULL;
GO
IF COL_LENGTH(N'dbo.TrabajadorHYT', N'FrecuenciaPago') IS NULL
    ALTER TABLE dbo.TrabajadorHYT ADD FrecuenciaPago VARCHAR(10) NULL;
GO
IF OBJECT_ID(N'dbo.CK_TrabajadorHYT_Planificacion', N'C') IS NULL
    ALTER TABLE dbo.TrabajadorHYT WITH CHECK ADD CONSTRAINT CK_TrabajadorHYT_Planificacion CHECK (
        ((HoraEntradaHabitual IS NULL AND HoraSalidaHabitual IS NULL)
         OR (HoraEntradaHabitual IS NOT NULL AND HoraSalidaHabitual IS NOT NULL AND HoraEntradaHabitual <> HoraSalidaHabitual))
        AND (RegimenPrevisional IS NULL OR RegimenPrevisional IN ('AFP', 'ONP'))
        AND (AdministradoraAFP IS NULL OR (RegimenPrevisional IS NOT NULL AND RegimenPrevisional = 'AFP'))
        AND (FrecuenciaPago IS NULL OR FrecuenciaPago IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')));
GO

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta trabajadores con cargo, sueldo, horario, regimen y frecuencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Tipo VARCHAR(10) = NULL,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar trabajadores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT trabajador.TrabajadorId, trabajador.Documento, trabajador.NombreCompleto, trabajador.Tipo,
               trabajador.FechaIngreso, trabajador.Banco, trabajador.CargoId, cargo.Nombre AS CargoNombre,
               CASE WHEN @TrabajadorId IS NOT NULL THEN trabajador.SueldoMensual END AS SueldoMensual,
               CASE WHEN trabajador.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(trabajador.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta = 1
                    THEN trabajador.NumeroCuenta END AS NumeroCuenta,
               trabajador.HoraEntradaHabitual, trabajador.HoraSalidaHabitual, trabajador.RegimenPrevisional,
               trabajador.AdministradoraAFP, trabajador.FrecuenciaPago, trabajador.Activo, trabajador.VersionFila, COUNT_BIG(1) OVER () AS Total
        FROM dbo.TrabajadorHYT AS trabajador
        LEFT JOIN dbo.CargoTrabajadorHYT AS cargo ON cargo.CargoId = trabajador.CargoId
        WHERE (@TrabajadorId IS NULL OR trabajador.TrabajadorId = @TrabajadorId)
          AND (@Tipo IS NULL OR trabajador.Tipo = @Tipo)
          AND (@SoloActivos = 0 OR trabajador.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, trabajador.NombreCompleto) > 0
               OR CHARINDEX(@Buscar, trabajador.Documento) > 0)
        ORDER BY trabajador.NombreCompleto, trabajador.TrabajadorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda trabajadores con cargo, sueldo base, horario, regimen y frecuencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Documento NVARCHAR(20),
    @NombreCompleto NVARCHAR(180),
    @Tipo VARCHAR(10),
    @FechaIngreso DATE,
    @Activo BIT,
    @TrabajadorId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ActualizarCuenta BIT = 0,
    @Banco NVARCHAR(80) = NULL,
    @NumeroCuenta NVARCHAR(34) = NULL,
    @CargoId BIGINT = NULL,
    @SueldoMensual DECIMAL(20,4) = NULL,
    @ActualizarPuesto BIT = 0,
    @HoraEntradaHabitual TIME(0) = NULL,
    @HoraSalidaHabitual TIME(0) = NULL,
    @RegimenPrevisional VARCHAR(3) = NULL,
    @AdministradoraAFP NVARCHAR(100) = NULL,
    @FrecuenciaPago VARCHAR(10) = NULL,
    @ActualizarPlanificacion BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar trabajadores.', 1;

        SET @NombreCompleto = NULLIF(TRIM(@NombreCompleto), N'');
        SET @Documento = NULLIF(TRIM(@Documento), N'');
        IF @Documento IS NULL OR @Tipo IS NULL OR @Tipo NOT IN ('PLANILLA', 'EXTERNO')
           OR @FechaIngreso IS NULL
           OR @NombreCompleto IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique documento, nombre, tipo y fecha de ingreso.', 1;

        SET @Banco = NULLIF(TRIM(@Banco), N'');
        SET @NumeroCuenta = NULLIF(TRIM(@NumeroCuenta), N'');
        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;
        END;

        SET @AdministradoraAFP = NULLIF(TRIM(@AdministradoraAFP), N'');
        IF @ActualizarPlanificacion = 1 AND (
            (@HoraEntradaHabitual IS NULL AND @HoraSalidaHabitual IS NOT NULL)
            OR (@HoraEntradaHabitual IS NOT NULL AND @HoraSalidaHabitual IS NULL)
            OR @HoraEntradaHabitual = @HoraSalidaHabitual
            OR (@RegimenPrevisional IS NOT NULL AND @RegimenPrevisional NOT IN ('AFP', 'ONP'))
            OR (@AdministradoraAFP IS NOT NULL AND COALESCE(@RegimenPrevisional, '') <> 'AFP')
            OR (@FrecuenciaPago IS NOT NULL AND @FrecuenciaPago NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')))
            THROW 51000, N'Revise el horario, regimen previsional y frecuencia de pago.', 1;

        BEGIN TRANSACTION;
        IF @ActualizarPuesto = 1
        BEGIN
            IF @SueldoMensual IS NOT NULL AND (@SueldoMensual <= 0 OR @SueldoMensual <> ROUND(@SueldoMensual, 2))
                THROW 51000, N'Indique un sueldo positivo con hasta dos decimales.', 1;
            IF @CargoId IS NOT NULL AND NOT EXISTS (
                SELECT 1 FROM dbo.CargoTrabajadorHYT AS cargo WITH (HOLDLOCK)
                WHERE cargo.CargoId = @CargoId AND (cargo.Activo = 1 OR EXISTS (
                    SELECT 1 FROM dbo.TrabajadorHYT AS trabajador
                    WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.CargoId = @CargoId)))
                THROW 51000, N'Seleccione un cargo activo.', 1;
        END;

        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @TrabajadorId IS NULL
        BEGIN
            INSERT INTO dbo.TrabajadorHYT (Documento, NombreCompleto, Tipo, FechaIngreso, Banco, NumeroCuenta, Activo, CargoId, SueldoMensual,
                 HoraEntradaHabitual, HoraSalidaHabitual, RegimenPrevisional, AdministradoraAFP, FrecuenciaPago)
            VALUES (@Documento, @NombreCompleto, @Tipo, @FechaIngreso, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END, @Activo,
                    CASE WHEN @ActualizarPuesto = 1 THEN @CargoId END, CASE WHEN @ActualizarPuesto = 1 THEN @SueldoMensual END,
                    CASE WHEN @ActualizarPlanificacion = 1 THEN @HoraEntradaHabitual END,
                    CASE WHEN @ActualizarPlanificacion = 1 THEN @HoraSalidaHabitual END,
                    CASE WHEN @ActualizarPlanificacion = 1 THEN @RegimenPrevisional END,
                    CASE WHEN @ActualizarPlanificacion = 1 THEN @AdministradoraAFP END,
                    CASE WHEN @ActualizarPlanificacion = 1 THEN @FrecuenciaPago END);
            SET @TrabajadorId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.TrabajadorHYT
            SET Documento = @Documento, NombreCompleto = @NombreCompleto, Activo = @Activo,
                Tipo = @Tipo, FechaIngreso = @FechaIngreso,
                CargoId = CASE WHEN @ActualizarPuesto = 1 THEN @CargoId ELSE CargoId END,
                SueldoMensual = CASE WHEN @ActualizarPuesto = 1 THEN @SueldoMensual ELSE SueldoMensual END,
                HoraEntradaHabitual = CASE WHEN @ActualizarPlanificacion = 1 THEN @HoraEntradaHabitual ELSE HoraEntradaHabitual END,
                HoraSalidaHabitual = CASE WHEN @ActualizarPlanificacion = 1 THEN @HoraSalidaHabitual ELSE HoraSalidaHabitual END,
                RegimenPrevisional = CASE WHEN @ActualizarPlanificacion = 1 THEN @RegimenPrevisional ELSE RegimenPrevisional END,
                AdministradoraAFP = CASE WHEN @ActualizarPlanificacion = 1 THEN @AdministradoraAFP ELSE AdministradoraAFP END,
                FrecuenciaPago = CASE WHEN @ActualizarPlanificacion = 1 THEN @FrecuenciaPago ELSE FrecuenciaPago END,
                Banco = CASE WHEN @ActualizarCuenta = 1 THEN @Banco ELSE Banco END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @TrabajadorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Documento AS Documento, @NombreCompleto AS NombreCompleto, @Activo AS Activo,
                   @Tipo AS Tipo, @FechaIngreso AS FechaIngreso, @ActualizarCuenta AS ActualizoCuenta,
                   @ActualizarPuesto AS ActualizoPuesto, @CargoId AS CargoId, @ActualizarPlanificacion AS ActualizoPlanificacion,
                   @HoraEntradaHabitual AS HoraEntradaHabitual, @HoraSalidaHabitual AS HoraSalidaHabitual,
                   @RegimenPrevisional AS RegimenPrevisional, @AdministradoraAFP AS AdministradoraAFP, @FrecuenciaPago AS FrecuenciaPago
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'TrabajadorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;

        SELECT @TrabajadorId AS TrabajadorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Tablas/PlanPagoHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PlanPagoHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PlanPagoHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PlanPagoHYT
    (
        PlanId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PlanPagoHYT PRIMARY KEY,
        Frecuencia VARCHAR(10) NOT NULL,
        Desde DATE NOT NULL,
        Hasta DATE NOT NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_PlanPagoHYT_Periodo UNIQUE (Frecuencia, Desde),
        CONSTRAINT CK_PlanPagoHYT_Rango CHECK (Desde <= Hasta AND Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL'))
    );
END;
GO

-- Fuente: Tablas/PlanPagoDetalleHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   PlanPagoDetalleHYT
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PlanPagoDetalleHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PlanPagoDetalleHYT
    (
        PlanId BIGINT NOT NULL CONSTRAINT FK_PlanPagoDetalleHYT_Plan REFERENCES dbo.PlanPagoHYT(PlanId),
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PlanPagoDetalleHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        Nombre NVARCHAR(180) NOT NULL,
        Tipo VARCHAR(10) NOT NULL,
        Base DECIMAL(18,2) NOT NULL,
        MinutosExtra INT NOT NULL,
        ImporteExtra DECIMAL(18,2) NOT NULL,
        Descuento DECIMAL(18,2) NOT NULL,
        CONSTRAINT PK_PlanPagoDetalleHYT PRIMARY KEY (PlanId, TrabajadorId),
        CONSTRAINT CK_PlanPagoDetalleHYT_Importes CHECK (Base >= 0 AND MinutosExtra >= 0
            AND ImporteExtra >= 0 AND Descuento >= 0 AND Descuento <= Base + ImporteExtra)
    );
END;
GO

-- Fuente: StoreProcedure/spu_PlanPagoHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_PlanPagoHYT_Consultar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar la planificacion.', 1;

        IF @Frecuencia IS NULL OR @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
           OR @Desde IS NULL OR @Desde > '99991224'
           OR (@Frecuencia = 'SEMANAL' AND DATEDIFF(DAY, '19000101', @Desde) % 7 <> 0)
           OR (@Frecuencia = 'QUINCENAL' AND DAY(@Desde) NOT IN (1,16))
           OR (@Frecuencia = 'MENSUAL' AND DAY(@Desde) <> 1)
            THROW 51000, N'Seleccione el inicio valido del periodo.', 1;

        DECLARE @Hasta DATE = CASE WHEN @Frecuencia = 'SEMANAL' THEN DATEADD(DAY, 6, @Desde)
            WHEN @Frecuencia = 'QUINCENAL' AND DAY(@Desde) = 1 THEN DATEADD(DAY, 14, @Desde)
            ELSE EOMONTH(@Desde) END;

        DECLARE @PlanId BIGINT, @Version BINARY(8);
        SELECT @PlanId = periodo.PlanId, @Version = periodo.VersionFila
        FROM dbo.PlanPagoHYT AS periodo WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;

        DECLARE @Filas NVARCHAR(MAX);
        IF @PlanId IS NOT NULL
            SET @Filas = (
                SELECT detalle.TrabajadorId, detalle.Nombre, detalle.Tipo, detalle.Base,
                       detalle.MinutosExtra, detalle.ImporteExtra, detalle.Descuento
                FROM dbo.PlanPagoDetalleHYT AS detalle WHERE detalle.PlanId = @PlanId
                ORDER BY detalle.Tipo, detalle.Nombre, detalle.TrabajadorId FOR JSON PATH);
        ELSE
            SET @Filas = (
                SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre, trabajador.Tipo,
                       COALESCE(trabajador.SueldoMensual, 0) AS Base, COALESCE(extras.Minutos, 0) AS MinutosExtra,
                       0 AS ImporteExtra, 0 AS Descuento
                FROM dbo.TrabajadorHYT AS trabajador
                OUTER APPLY (
                    SELECT SUM(jornada.MinutosExtrasAprobados) AS Minutos
                    FROM dbo.JornadaHYT AS jornada
                    WHERE jornada.TrabajadorId = trabajador.TrabajadorId
                      AND CONVERT(DATE, jornada.Entrada) BETWEEN @Desde AND @Hasta
                ) AS extras
                WHERE trabajador.Activo = 1 AND trabajador.FrecuenciaPago = @Frecuencia AND trabajador.FechaIngreso <= @Hasta
                ORDER BY trabajador.Tipo, trabajador.NombreCompleto, trabajador.TrabajadorId FOR JSON PATH);

        SELECT (SELECT @PlanId AS PlanId, @Version AS VersionFila, @Frecuencia AS Frecuencia,
            @Desde AS Desde, @Hasta AS Hasta, JSON_QUERY(COALESCE(@Filas, N'[]')) AS Filas
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PlanPagoHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   spu_PlanPagoHYT_Guardar
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE,
    @VersionFila BINARY(8),
    @Filas NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar la planificacion.', 1;

        IF @Frecuencia IS NULL OR @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
           OR @Desde IS NULL OR @Desde > '99991224'
           OR (@Frecuencia = 'SEMANAL' AND DATEDIFF(DAY, '19000101', @Desde) % 7 <> 0)
           OR (@Frecuencia = 'QUINCENAL' AND DAY(@Desde) NOT IN (1,16))
           OR (@Frecuencia = 'MENSUAL' AND DAY(@Desde) <> 1)
            THROW 51000, N'Seleccione el inicio valido del periodo.', 1;

        DECLARE @Hasta DATE = CASE WHEN @Frecuencia = 'SEMANAL' THEN DATEADD(DAY, 6, @Desde)
            WHEN @Frecuencia = 'QUINCENAL' AND DAY(@Desde) = 1 THEN DATEADD(DAY, 14, @Desde)
            ELSE EOMONTH(@Desde) END;

        IF COALESCE(ISJSON(@Filas, ARRAY), 0) <> 1
            THROW 51000, N'La planificacion no es valida.', 1;

        DECLARE @Datos TABLE (TrabajadorId BIGINT PRIMARY KEY, Base DECIMAL(20,4), MinutosExtra INT,
                              ImporteExtra DECIMAL(20,4), Descuento DECIMAL(20,4));
        INSERT INTO @Datos (TrabajadorId, Base, MinutosExtra, ImporteExtra, Descuento)
        SELECT fila.TrabajadorId, fila.Base, fila.MinutosExtra, fila.ImporteExtra, fila.Descuento
        FROM OPENJSON(@Filas) WITH (TrabajadorId BIGINT, Base DECIMAL(20,4), MinutosExtra INT,
            ImporteExtra DECIMAL(20,4), Descuento DECIMAL(20,4)) AS fila;
        IF NOT EXISTS (SELECT 1 FROM @Datos) OR EXISTS (
            SELECT 1 FROM @Datos AS fila
            WHERE fila.Base IS NULL OR fila.ImporteExtra IS NULL OR fila.Descuento IS NULL OR fila.MinutosExtra IS NULL
               OR fila.Base < 0 OR fila.ImporteExtra < 0 OR fila.Descuento < 0 OR fila.MinutosExtra < 0
               OR fila.Descuento > fila.Base + fila.ImporteExtra
               OR fila.Base <> ROUND(fila.Base, 2) OR fila.ImporteExtra <> ROUND(fila.ImporteExtra, 2)
               OR fila.Descuento <> ROUND(fila.Descuento, 2))
            THROW 51000, N'Revise los importes y minutos de la planificacion.', 1;

        BEGIN TRANSACTION;
        DECLARE @PlanId BIGINT, @Actual BINARY(8);
        SELECT @PlanId = periodo.PlanId, @Actual = periodo.VersionFila
        FROM dbo.PlanPagoHYT AS periodo WITH (UPDLOCK, HOLDLOCK)
        WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;
        IF @Actual IS DISTINCT FROM @VersionFila
            THROW 51000, N'La planificacion cambio. Vuelva a abrirla.', 1;

        IF @PlanId IS NULL
        BEGIN
            IF EXISTS (
                SELECT 1 FROM @Datos AS fila
                LEFT JOIN dbo.TrabajadorHYT AS trabajador WITH (HOLDLOCK) ON trabajador.TrabajadorId = fila.TrabajadorId
                WHERE trabajador.TrabajadorId IS NULL OR trabajador.Activo = 0
                   OR trabajador.FrecuenciaPago IS DISTINCT FROM @Frecuencia OR trabajador.FechaIngreso > @Hasta)
                THROW 51000, N'El personal del periodo cambio. Vuelva a consultarlo.', 1;

            INSERT INTO dbo.PlanPagoHYT (Frecuencia, Desde, Hasta) VALUES (@Frecuencia, @Desde, @Hasta);
            SET @PlanId = SCOPE_IDENTITY();
            INSERT INTO dbo.PlanPagoDetalleHYT (PlanId, TrabajadorId, Nombre, Tipo, Base, MinutosExtra, ImporteExtra, Descuento)
            SELECT @PlanId, fila.TrabajadorId, trabajador.NombreCompleto, trabajador.Tipo,
                   fila.Base, fila.MinutosExtra, fila.ImporteExtra, fila.Descuento
            FROM @Datos AS fila INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = fila.TrabajadorId;
        END
        ELSE
        BEGIN
            IF EXISTS (SELECT fila.TrabajadorId FROM @Datos AS fila EXCEPT
                       SELECT detalle.TrabajadorId FROM dbo.PlanPagoDetalleHYT AS detalle WHERE detalle.PlanId = @PlanId)
               OR EXISTS (SELECT detalle.TrabajadorId FROM dbo.PlanPagoDetalleHYT AS detalle WHERE detalle.PlanId = @PlanId EXCEPT
                          SELECT fila.TrabajadorId FROM @Datos AS fila)
                THROW 51000, N'El detalle no coincide con la planificacion guardada.', 1;

            UPDATE detalle SET Base = fila.Base, MinutosExtra = fila.MinutosExtra,
                ImporteExtra = fila.ImporteExtra, Descuento = fila.Descuento
            FROM dbo.PlanPagoDetalleHYT AS detalle INNER JOIN @Datos AS fila ON fila.TrabajadorId = detalle.TrabajadorId
            WHERE detalle.PlanId = @PlanId;
            UPDATE dbo.PlanPagoHYT SET Hasta = @Hasta WHERE PlanId = @PlanId;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PlanId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PlanPagoHYT', @RegistroId, N'GUARDAR';
        COMMIT TRANSACTION;
        SELECT @PlanId AS PlanId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/040_PermisosPlanificacionPago.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Concede acceso a planificacion con los permisos de pagos de personal.
-- =============================================
USE [SistemaWebHYT];
GO
GRANT EXECUTE ON OBJECT::dbo.spu_PlanPagoHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PlanPagoHYT_Guardar TO hyt_app_executor;
GO

COMMIT TRANSACTION;
GO
