-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda trabajadores con cargo, sueldo base, horario, regimen y frecuencia. Incluye el CCI con permiso bancario.
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
    @ActualizarPlanificacion BIT = 0,
    @CuentaInterbancaria VARCHAR(100) = NULL,
    @ActualizarCCI BIT = 0
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
        SET @CuentaInterbancaria = NULLIF(TRIM(@CuentaInterbancaria), '');
        IF @ActualizarCCI = 1 AND @ActualizarCuenta <> 1
            THROW 51000, N'El CCI requiere autorizacion para actualizar la cuenta bancaria.', 1;

        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;

            IF @ActualizarCCI = 1 AND @CuentaInterbancaria IS NOT NULL
                AND (@Banco IS NULL OR DATALENGTH(@CuentaInterbancaria) <> 20
                     OR @CuentaInterbancaria COLLATE Latin1_General_100_BIN2 LIKE '%[^0-9]%')
                THROW 51000, N'Indique un banco y un CCI de veinte digitos.', 1;
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
            INSERT INTO dbo.TrabajadorHYT (Documento, NombreCompleto, Tipo, FechaIngreso, Banco, NumeroCuenta, CuentaInterbancaria, Activo, CargoId, SueldoMensual,
                 HoraEntradaHabitual, HoraSalidaHabitual, RegimenPrevisional, AdministradoraAFP, FrecuenciaPago)
            VALUES (@Documento, @NombreCompleto, @Tipo, @FechaIngreso, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END,
                    CASE WHEN @ActualizarCuenta = 1 AND @ActualizarCCI = 1 THEN @CuentaInterbancaria END, @Activo,
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
                CuentaInterbancaria = CASE WHEN @ActualizarCuenta = 1 AND @Banco IS NULL THEN NULL
                    WHEN @ActualizarCuenta = 1 AND @ActualizarCCI = 1 THEN @CuentaInterbancaria ELSE CuentaInterbancaria END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @TrabajadorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Documento AS Documento, @NombreCompleto AS NombreCompleto, @Activo AS Activo,
                   @Tipo AS Tipo, @FechaIngreso AS FechaIngreso, @ActualizarCuenta AS ActualizoCuenta, @ActualizarCCI AS ActualizoCCI,
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
