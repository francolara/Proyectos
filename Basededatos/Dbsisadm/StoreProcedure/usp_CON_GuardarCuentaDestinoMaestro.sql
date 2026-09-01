-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Crea o actualiza una regla maestra y sus tramos usando codigos contables portables.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarCuentaDestinoMaestro
    @IdCuentaDestinoReglaMaestro INT = NULL,
    @CodigoCuentaOrigen VARCHAR(20),
    @Activo BIT,
    @Observacion NVARCHAR(500) = NULL,
    @DetallesJson NVARCHAR(MAX),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY
        DECLARE @CodigoOrigen VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CodigoCuentaOrigen)), '');

        DECLARE @Detalles TABLE
        (
            Orden SMALLINT NOT NULL,
            CodigoCuentaDestinoCargo VARCHAR(20) NOT NULL,
            CodigoCuentaDestinoAbono VARCHAR(20) NOT NULL,
            Porcentaje DECIMAL(7,4) NOT NULL,
            Activo BIT NOT NULL
        );

        INSERT INTO @Detalles (Orden, CodigoCuentaDestinoCargo, CodigoCuentaDestinoAbono, Porcentaje, Activo)
        SELECT Orden, LTRIM(RTRIM(CodigoCuentaDestinoCargo)), LTRIM(RTRIM(CodigoCuentaDestinoAbono)), Porcentaje, Activo
        FROM OPENJSON(@DetallesJson)
        WITH
        (
            Orden SMALLINT '$.orden',
            CodigoCuentaDestinoCargo VARCHAR(20) '$.codigoCuentaDestinoCargo',
            CodigoCuentaDestinoAbono VARCHAR(20) '$.codigoCuentaDestinoAbono',
            Porcentaje DECIMAL(7,4) '$.porcentaje',
            Activo BIT '$.activo'
        );

        IF @CodigoOrigen IS NULL
            RAISERROR(N'La cuenta origen es obligatoria.', 16, 1);

        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.CON_PlanCuentaMaestro
            WHERE CodigoCuenta = @CodigoOrigen AND Estado = 1 AND AceptaMovimiento = 1
        )
            RAISERROR(N'La cuenta origen no existe, esta inactiva o no acepta movimiento.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM @Detalles WHERE Activo = 1)
            RAISERROR(N'Debe registrar al menos un tramo activo.', 16, 1);

        IF EXISTS (SELECT Orden FROM @Detalles GROUP BY Orden HAVING COUNT(1) > 1)
            RAISERROR(N'El orden de los tramos no puede repetirse.', 16, 1);

        IF (SELECT ISNULL(SUM(Porcentaje), 0) FROM @Detalles WHERE Activo = 1) <> 100
            RAISERROR(N'La suma de los porcentajes activos debe ser exactamente 100.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS detalle
            CROSS APPLY
            (
                VALUES (detalle.CodigoCuentaDestinoCargo), (detalle.CodigoCuentaDestinoAbono)
            ) AS codigo (CodigoCuenta)
            LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
                ON cuenta.CodigoCuenta = codigo.CodigoCuenta
               AND cuenta.Estado = 1
               AND cuenta.AceptaMovimiento = 1
            WHERE detalle.Activo = 1
              AND cuenta.IdPlanCuentaMaestro IS NULL
        )
            RAISERROR(N'Una cuenta destino no existe, esta inactiva o no acepta movimiento.', 16, 1);

        IF EXISTS
        (
            SELECT 1 FROM dbo.CON_CuentaDestinoReglaMaestro
            WHERE CodigoCuentaOrigen = @CodigoOrigen
              AND (@IdCuentaDestinoReglaMaestro IS NULL OR IdCuentaDestinoReglaMaestro <> @IdCuentaDestinoReglaMaestro)
        )
            RAISERROR(N'Ya existe una regla maestra para la cuenta origen indicada.', 16, 1);

        IF @IdCuentaDestinoReglaMaestro IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.CON_CuentaDestinoReglaMaestro
               WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro
                 AND CodigoCuentaOrigen <> @CodigoOrigen
           )
            RAISERROR(N'La cuenta origen no puede modificarse despues de crear la regla maestra.', 16, 1);

        BEGIN TRANSACTION;

        IF @IdCuentaDestinoReglaMaestro IS NULL
        BEGIN
            INSERT INTO dbo.CON_CuentaDestinoReglaMaestro (CodigoCuentaOrigen, Activo, Observacion, UsuarioRegistro)
            VALUES (@CodigoOrigen, @Activo, NULLIF(LTRIM(RTRIM(@Observacion)), N''), @UsuarioRegistro);
            SET @IdCuentaDestinoReglaMaestro = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.CON_CuentaDestinoReglaMaestro WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro)
                RAISERROR(N'La regla maestra indicada no existe.', 16, 1);

            UPDATE dbo.CON_CuentaDestinoReglaMaestro
            SET CodigoCuentaOrigen = @CodigoOrigen,
                Activo = @Activo,
                Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), N''),
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro;

            DELETE FROM dbo.CON_CuentaDestinoReglaDetalleMaestro
            WHERE IdCuentaDestinoReglaMaestro = @IdCuentaDestinoReglaMaestro;
        END;

        INSERT INTO dbo.CON_CuentaDestinoReglaDetalleMaestro
        (
            IdCuentaDestinoReglaMaestro, Orden, CodigoCuentaDestinoCargo,
            CodigoCuentaDestinoAbono, Porcentaje, Activo, UsuarioRegistro
        )
        SELECT @IdCuentaDestinoReglaMaestro, Orden, CodigoCuentaDestinoCargo,
               CodigoCuentaDestinoAbono, Porcentaje, Activo, @UsuarioRegistro
        FROM @Detalles;

        COMMIT TRANSACTION;
        SELECT @IdCuentaDestinoReglaMaestro AS IdCuentaDestinoReglaMaestro;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
