-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Registra una transferencia entre cuentas generando los movimientos bancarios emisor/receptor enlazados entre si y sus asientos contables automaticos, usando la cuenta corriente contraria en la glosa detalle y el nro de operacion en Referencia.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Ajusta el detalle de transferencias para mostrar la cuenta corriente contraria en la glosa y mover el nro de operacion a Referencia, evitando que aparezca en RUC/DNI.
-- Firma: FRANCO LARA - 09/07/2026 | Obtiene y guarda tipo de cambio independiente por fecha en emisor/receptor, elimina la restriccion de igualdad entre ambos y permite persistir el importe real recibido en la cuenta receptora cuando las monedas difieren.

CREATE OR ALTER PROCEDURE dbo.usp_BAN_GuardarTransferenciaCuenta
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresaEmisor INT,
    @IdBancoConfiguracionEmpresaReceptor INT,
    @IdOpeBancariaEmisor CHAR(2),
    @IdOpeBancariaReceptor CHAR(2),
    @FechaEmisionEmisor DATE,
    @FechaEmisionReceptor DATE,
    @TipoCambioEmisor DECIMAL(18, 6),
    @TipoCambioReceptor DECIMAL(18, 6),
    @NumeroOperacionEmisor VARCHAR(20) = NULL,
    @NumeroOperacionReceptor VARCHAR(20) = NULL,
    @ImporteEmisor DECIMAL(18, 2),
    @ImporteReceptor DECIMAL(18, 2),
    @GlosaEmisor NVARCHAR(300),
    @GlosaReceptor NVARCHAR(300),
    @ObservacionEmisor NVARCHAR(500) = NULL,
    @ObservacionReceptor NVARCHAR(500) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdTransferenciaCuenta UNIQUEIDENTIFIER = NEWID();
        DECLARE @IdMovimientoBancoEmisor INT = NULL;
        DECLARE @IdMovimientoBancoReceptor INT = NULL;
        DECLARE @IdAsientoEmisor INT = NULL;
        DECLARE @IdAsientoReceptor INT = NULL;
        DECLARE @NumeroMovimientoEmisor INT = NULL;
        DECLARE @NumeroMovimientoReceptor INT = NULL;
        DECLARE @NumeroAsientoEmisor INT = NULL;
        DECLARE @NumeroAsientoReceptor INT = NULL;
        DECLARE @IdPlanCuentaEmisor INT = NULL;
        DECLARE @IdPlanCuentaReceptor INT = NULL;
        DECLARE @NroCuentaCorrienteEmisor VARCHAR(50) = NULL;
        DECLARE @NroCuentaCorrienteReceptor VARCHAR(50) = NULL;
        DECLARE @CodigoMonedaEmisor VARCHAR(10) = NULL;
        DECLARE @CodigoMonedaReceptor VARCHAR(10) = NULL;
        DECLARE @DetallesXmlEmisor XML;
        DECLARE @DetallesXmlReceptor XML;

        IF @IdBancoConfiguracionEmpresaEmisor = @IdBancoConfiguracionEmpresaReceptor
        BEGIN
            RAISERROR('La cuenta corriente emisora debe ser distinta de la receptora.', 16, 1);
        END;

        IF @ImporteEmisor <= 0
        BEGIN
            RAISERROR('Ingrese un monto mayor a cero para la transferencia.', 16, 1);
        END;

        IF @TipoCambioEmisor <= 0 OR @TipoCambioReceptor <= 0
        BEGIN
            RAISERROR('El tipo de cambio debe ser mayor a cero en ambas secciones.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.operacionesbancarias AS op
            WHERE LTRIM(RTRIM(op.idOpeBancaria)) = LTRIM(RTRIM(@IdOpeBancariaEmisor))
              AND LTRIM(RTRIM(op.Destino)) = 'E'
              AND LTRIM(RTRIM(op.idTipoOpeBancaria)) = 'T'
        )
        BEGIN
            RAISERROR('La operacion bancaria del emisor no corresponde a una transferencia de egreso.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.operacionesbancarias AS op
            WHERE LTRIM(RTRIM(op.idOpeBancaria)) = LTRIM(RTRIM(@IdOpeBancariaReceptor))
              AND LTRIM(RTRIM(op.Destino)) = 'I'
              AND LTRIM(RTRIM(op.idTipoOpeBancaria)) = 'T'
        )
        BEGIN
            RAISERROR('La operacion bancaria del receptor no corresponde a una transferencia de ingreso.', 16, 1);
        END;

        SELECT
            @IdPlanCuentaEmisor = cc.IdPlanCuenta,
            @NroCuentaCorrienteEmisor = cc.NroCuentaCorriente,
            @CodigoMonedaEmisor = UPPER(LTRIM(RTRIM(ISNULL(mon.CodigoMoneda, ''))))
        FROM dbo.CON_BancosConfiguracionEmpresa AS cc
        LEFT JOIN dbo.ADM_Moneda AS mon
            ON mon.IdMoneda = cc.IdMoneda
        WHERE cc.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresaEmisor
          AND cc.IdEmpresa = @IdEmpresa
          AND cc.Activo = 1;

        SELECT
            @IdPlanCuentaReceptor = cc.IdPlanCuenta,
            @NroCuentaCorrienteReceptor = cc.NroCuentaCorriente,
            @CodigoMonedaReceptor = UPPER(LTRIM(RTRIM(ISNULL(mon.CodigoMoneda, ''))))
        FROM dbo.CON_BancosConfiguracionEmpresa AS cc
        LEFT JOIN dbo.ADM_Moneda AS mon
            ON mon.IdMoneda = cc.IdMoneda
        WHERE cc.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresaReceptor
          AND cc.IdEmpresa = @IdEmpresa
          AND cc.Activo = 1;

        IF @IdPlanCuentaEmisor IS NULL OR @IdPlanCuentaReceptor IS NULL
        BEGIN
            RAISERROR('Las cuentas corrientes de la transferencia deben existir y estar activas para la empresa.', 16, 1);
        END;

        IF NULLIF(@CodigoMonedaEmisor, '') IS NULL OR NULLIF(@CodigoMonedaReceptor, '') IS NULL
        BEGIN
            RAISERROR('Ambas cuentas corrientes deben tener una moneda configurada.', 16, 1);
        END;

        IF @CodigoMonedaEmisor = @CodigoMonedaReceptor
        BEGIN
            SET @ImporteReceptor = @ImporteEmisor;
        END;
        ELSE IF @ImporteReceptor <= 0
        BEGIN
            IF @CodigoMonedaEmisor = 'USD' AND @CodigoMonedaReceptor = 'PEN'
            BEGIN
                SET @ImporteReceptor = ROUND(@ImporteEmisor * @TipoCambioEmisor, 2);
            END;
            ELSE IF @CodigoMonedaEmisor = 'PEN' AND @CodigoMonedaReceptor = 'USD'
            BEGIN
                SET @ImporteReceptor = ROUND(@ImporteEmisor / @TipoCambioEmisor, 2);
            END;
            ELSE
            BEGIN
                RAISERROR('Solo se admite conversion automatica entre cuentas en PEN y USD.', 16, 1);
            END;
        END;
        ELSE IF @CodigoMonedaEmisor NOT IN ('PEN', 'USD') OR @CodigoMonedaReceptor NOT IN ('PEN', 'USD')
        BEGIN
            RAISERROR('Solo se admite conversion automatica entre cuentas en PEN y USD.', 16, 1);
        END;

        SET @DetallesXmlEmisor =
        (
            SELECT
                1 AS [@Item],
                @IdPlanCuentaReceptor AS [@IdPlanCuenta],
                LEFT(CONCAT(N'Banco ', ISNULL(@NroCuentaCorrienteReceptor, N'')), 300) AS [@GlosaDetalle],
                NULL AS [@NumeroDocumento],
                NULLIF(LTRIM(RTRIM(@NumeroOperacionEmisor)), '') AS [@ReferenciaLinea],
                @TipoCambioEmisor AS [@TipoCambioLinea],
                @ImporteEmisor AS [@Debe],
                CAST(0 AS DECIMAL(18,2)) AS [@Haber]
            FOR XML PATH('Detalle'), ROOT('Detalles'), TYPE
        );

        SET @DetallesXmlReceptor =
        (
            SELECT
                1 AS [@Item],
                @IdPlanCuentaEmisor AS [@IdPlanCuenta],
                LEFT(CONCAT(N'Banco ', ISNULL(@NroCuentaCorrienteEmisor, N'')), 300) AS [@GlosaDetalle],
                NULL AS [@NumeroDocumento],
                NULLIF(LTRIM(RTRIM(@NumeroOperacionReceptor)), '') AS [@ReferenciaLinea],
                @TipoCambioReceptor AS [@TipoCambioLinea],
                CAST(0 AS DECIMAL(18,2)) AS [@Debe],
                @ImporteReceptor AS [@Haber]
            FOR XML PATH('Detalle'), ROOT('Detalles'), TYPE
        );

        BEGIN TRANSACTION;

        EXEC dbo.usp_BAN_GuardarMovimientoBanco
            @IdMovimientoBanco = NULL,
            @IdEmpresa = @IdEmpresa,
            @IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresaEmisor,
            @TipoMovimiento = 'E',
            @IdOpeBancaria = @IdOpeBancariaEmisor,
            @FechaEmision = @FechaEmisionEmisor,
            @TipoCambio = @TipoCambioEmisor,
            @IdPersona = NULL,
            @NumeroDocumento = @NumeroOperacionEmisor,
            @Glosa = @GlosaEmisor,
            @Observacion = @ObservacionEmisor,
            @ImporteTotal = @ImporteEmisor,
            @UsuarioRegistro = @UsuarioRegistro,
            @DetallesXml = @DetallesXmlEmisor,
            @IdTransferenciaCuenta = @IdTransferenciaCuenta,
            @RolTransferencia = 'E',
            @IdMovimientoBancoRelacionado = NULL,
            @RetornarResultado = 0,
            @IdMovimientoBancoGenerado = @IdMovimientoBancoEmisor OUTPUT,
            @IdAsientoGenerado = @IdAsientoEmisor OUTPUT,
            @NumeroMovimientoGenerado = @NumeroMovimientoEmisor OUTPUT,
            @NumeroAsientoGenerado = @NumeroAsientoEmisor OUTPUT;

        EXEC dbo.usp_BAN_GuardarMovimientoBanco
            @IdMovimientoBanco = NULL,
            @IdEmpresa = @IdEmpresa,
            @IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresaReceptor,
            @TipoMovimiento = 'I',
            @IdOpeBancaria = @IdOpeBancariaReceptor,
            @FechaEmision = @FechaEmisionReceptor,
            @TipoCambio = @TipoCambioReceptor,
            @IdPersona = NULL,
            @NumeroDocumento = @NumeroOperacionReceptor,
            @Glosa = @GlosaReceptor,
            @Observacion = @ObservacionReceptor,
            @ImporteTotal = @ImporteReceptor,
            @UsuarioRegistro = @UsuarioRegistro,
            @DetallesXml = @DetallesXmlReceptor,
            @IdTransferenciaCuenta = @IdTransferenciaCuenta,
            @RolTransferencia = 'I',
            @IdMovimientoBancoRelacionado = @IdMovimientoBancoEmisor,
            @RetornarResultado = 0,
            @IdMovimientoBancoGenerado = @IdMovimientoBancoReceptor OUTPUT,
            @IdAsientoGenerado = @IdAsientoReceptor OUTPUT,
            @NumeroMovimientoGenerado = @NumeroMovimientoReceptor OUTPUT,
            @NumeroAsientoGenerado = @NumeroAsientoReceptor OUTPUT;

        UPDATE dbo.BAN_MovimientoBanco
        SET IdMovimientoBancoRelacionado = @IdMovimientoBancoReceptor
        WHERE IdMovimientoBanco = @IdMovimientoBancoEmisor
          AND IdEmpresa = @IdEmpresa;

        COMMIT TRANSACTION;

        SELECT
            @IdTransferenciaCuenta AS IdTransferenciaCuenta,
            @IdMovimientoBancoEmisor AS IdMovimientoBancoEmisor,
            @NumeroMovimientoEmisor AS NumeroMovimientoEmisor,
            @NumeroAsientoEmisor AS NumeroAsientoEmisor,
            @IdMovimientoBancoReceptor AS IdMovimientoBancoReceptor,
            @NumeroMovimientoReceptor AS NumeroMovimientoReceptor,
            @NumeroAsientoReceptor AS NumeroAsientoReceptor,
            @ImporteEmisor AS ImporteEmisor,
            @ImporteReceptor AS ImporteReceptor;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH

END
