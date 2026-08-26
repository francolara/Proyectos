-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Registra o reemplaza la configuracion de cuentas destino para una cuenta origen, empresa y ejercicio.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Guarda una sola configuracion de cuentas destino por empresa y cuenta origen, sin depender de un ejercicio.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Elimina Ejercicio del guardado porque la regla es unica por empresa y cuenta origen.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarCuentaDestinoRegla
    @IdEmpresa INT,
    @IdPlanCuentaOrigen INT,
    @Activo BIT,
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaDestinoRegla INT
        DECLARE @PorcentajeTotal DECIMAL(18,4)

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de cuentas destino.', 16, 1);
        END;

        DECLARE @Detalle TABLE
        (
            Orden SMALLINT NOT NULL,
            IdPlanCuentaDestinoCargo INT NOT NULL,
            IdPlanCuentaDestinoAbono INT NOT NULL,
            Porcentaje DECIMAL(7,4) NOT NULL,
            Activo BIT NOT NULL
        );

        INSERT INTO @Detalle
        (
            Orden,
            IdPlanCuentaDestinoCargo,
            IdPlanCuentaDestinoAbono,
            Porcentaje,
            Activo
        )
        SELECT
            T.N.value('@Orden', 'smallint'),
            T.N.value('@IdPlanCuentaDestinoCargo', 'int'),
            T.N.value('@IdPlanCuentaDestinoAbono', 'int'),
            T.N.value('@Porcentaje', 'decimal(7,4)'),
            T.N.value('@Activo', 'bit')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @Detalle
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos un tramo de cuenta destino.', 16, 1);
        END;

        SELECT
            @PorcentajeTotal = SUM(d.Porcentaje)
        FROM @Detalle AS d
        WHERE d.Activo = 1;

        IF ISNULL(@PorcentajeTotal, 0) <> 100
        BEGIN
            RAISERROR(N'La suma del porcentaje de los destinos activos debe ser 100.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            WHERE d.IdPlanCuentaDestinoCargo = d.IdPlanCuentaDestinoAbono
        )
        BEGIN
            RAISERROR(N'La cuenta destino cargo y abono no pueden ser iguales en el mismo tramo.', 16, 1);
        END;

        BEGIN TRAN;

        SELECT
            @IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
        FROM dbo.CON_CuentaDestinoRegla AS r
        WHERE r.IdEmpresa = @IdEmpresa
          AND r.IdPlanCuentaOrigen = @IdPlanCuentaOrigen;

        IF @IdCuentaDestinoRegla IS NULL
        BEGIN
            INSERT INTO dbo.CON_CuentaDestinoRegla
            (
                IdEmpresa,
                IdPlanCuentaOrigen,
                Activo,
                Observacion,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdPlanCuentaOrigen,
                @Activo,
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdCuentaDestinoRegla = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.CON_CuentaDestinoRegla
            SET Activo = @Activo,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdCuentaDestinoRegla = @IdCuentaDestinoRegla;

            DELETE FROM dbo.CON_CuentaDestinoReglaDetalle
            WHERE IdCuentaDestinoRegla = @IdCuentaDestinoRegla;
        END;

        INSERT INTO dbo.CON_CuentaDestinoReglaDetalle
        (
            IdCuentaDestinoRegla,
            Orden,
            IdPlanCuentaDestinoCargo,
            IdPlanCuentaDestinoAbono,
            Porcentaje,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdCuentaDestinoRegla,
            d.Orden,
            d.IdPlanCuentaDestinoCargo,
            d.IdPlanCuentaDestinoAbono,
            d.Porcentaje,
            d.Activo,
            @UsuarioRegistro
        FROM @Detalle AS d
        ORDER BY
            d.Orden ASC;

        COMMIT;

        SELECT
            @IdCuentaDestinoRegla AS IdCuentaDestinoRegla;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
