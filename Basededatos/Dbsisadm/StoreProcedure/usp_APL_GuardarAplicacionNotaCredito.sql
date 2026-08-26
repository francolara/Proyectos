-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Registra una aplicacion entre un comprobante pendiente y una nota de credito, actualiza ambos saldos y genera asiento segun la provision APNC.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Agrega el guardado del modulo Aplicaciones con soporte de importe parcial, descuento de saldos en compras/ventas, tipo de cambio editable para el asiento y mensaje claro cuando falle la generacion contable.
-- Firma: FRANCO LARA - 25/06/2026 | Ajusta la generacion del asiento de Aplicaciones para dejar su estado final en PROVISIONADO en lugar de BORRADOR.
-- Firma: FRANCO LARA - 03/07/2026 | Incluye DH en el XML del asiento automatico de aplicaciones para propagar el sentido contable al guardado centralizado del detalle.
-- Firma: FRANCO LARA - 25/08/2026 | Exige las cuentas de comprobantes configuradas y activas por empresa, sin usar respaldos del maestro.

CREATE OR ALTER PROCEDURE dbo.usp_APL_GuardarAplicacionNotaCredito
    @IdEmpresa INT,
    @ModuloOperacion VARCHAR(10),
    @IdPersona INT,
    @FechaAplicacion DATE,
    @TipoCambioAplicacion DECIMAL(18, 6),
    @IdRegistroComprobante INT,
    @IdRegistroNotaCredito INT,
    @ImporteAplicado DECIMAL(18, 2),
    @Glosa NVARCHAR(300),
    @Observacion NVARCHAR(500) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdAplicacionNotaCredito INT = NULL;
        DECLARE @IdAsiento INT = NULL;
        DECLARE @NumeroAsiento INT = NULL;
        DECLARE @IdMoneda INT = NULL;
        DECLARE @IdMonedaNotaCredito INT = NULL;
        DECLARE @CodigoMoneda VARCHAR(10) = NULL;
        DECLARE @CodigoMonedaNotaCredito VARCHAR(10) = NULL;
        DECLARE @TipoCambioDocumento DECIMAL(18, 6) = 1;
        DECLARE @TipoCambioAsiento DECIMAL(18, 6) = 1;
        DECLARE @SaldoComprobante DECIMAL(18, 2) = 0;
        DECLARE @SaldoNotaCredito DECIMAL(18, 2) = 0;
        DECLARE @ImporteTotalComprobante DECIMAL(18, 2) = 0;
        DECLARE @ImporteTotalNotaCredito DECIMAL(18, 2) = 0;
        DECLARE @TipoComprobanteAplicado VARCHAR(3) = NULL;
        DECLARE @TipoComprobanteNc VARCHAR(3) = NULL;
        DECLARE @SerieAplicado VARCHAR(10) = NULL;
        DECLARE @SerieNc VARCHAR(10) = NULL;
        DECLARE @NumeroAplicado VARCHAR(20) = NULL;
        DECLARE @NumeroNc VARCHAR(20) = NULL;
        DECLARE @DescripcionTipoAplicado NVARCHAR(150) = NULL;
        DECLARE @DescripcionTipoNc NVARCHAR(150) = NULL;
        DECLARE @NumeroDocumentoPersona VARCHAR(20) = NULL;
        DECLARE @IdOrigen INT = NULL;
        DECLARE @GeneraAsientoAutomatico BIT = 0;
        DECLARE @ConfiguracionActiva BIT = 0;
        DECLARE @IdCuentaComprobante INT = NULL;
        DECLARE @IdCuentaNotaCredito INT = NULL;
        DECLARE @DetalleXml XML = NULL;
        DECLARE @GlosaAsiento NVARCHAR(500) = NULL;
        DECLARE @ReferenciaExterna NVARCHAR(100) = NULL;

        DECLARE @DetalleAsiento TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NOT NULL,
            TipoDocumento NVARCHAR(150) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            TipoCambioLinea DECIMAL(18, 6) NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL
        );

        DECLARE @ResultadoAsiento TABLE
        (
            IdAsiento INT NOT NULL,
            Periodo CHAR(6) NOT NULL,
            NumeroAsiento INT NOT NULL,
            TotalDebe DECIMAL(18, 2) NOT NULL,
            TotalHaber DECIMAL(18, 2) NOT NULL,
            Estado NVARCHAR(20) NOT NULL
        );

        IF @ModuloOperacion NOT IN ('COM', 'VEN')
        BEGIN
            RAISERROR(N'El modulo de aplicaciones debe ser COM o VEN.', 16, 1);
        END;

        IF @IdRegistroComprobante = @IdRegistroNotaCredito
        BEGIN
            RAISERROR(N'El comprobante y la nota de credito deben ser distintos.', 16, 1);
        END;

        IF @ImporteAplicado <= 0
        BEGIN
            RAISERROR(N'Ingrese un importe aplicado mayor a cero.', 16, 1);
        END;

        IF NULLIF(LTRIM(RTRIM(@Glosa)), N'') IS NULL
        BEGIN
            RAISERROR(N'Ingrese la glosa de la aplicacion.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Persona AS p
            WHERE p.IdPersona = @IdPersona
              AND p.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La persona indicada no pertenece a la empresa activa.', 16, 1);
        END;

        IF @ModuloOperacion = 'VEN'
        BEGIN
            SELECT
                @SaldoComprobante = v.Saldo,
                @ImporteTotalComprobante = v.ImporteTotal,
                @IdMoneda = v.IdMoneda,
                @TipoCambioDocumento = v.TipoCambio,
                @TipoComprobanteAplicado = v.TipoComprobante,
                @SerieAplicado = v.Serie,
                @NumeroAplicado = v.Numero,
                @NumeroDocumentoPersona = pe.NumeroDocumento,
                @DescripcionTipoAplicado = tc.Descripcion
            FROM dbo.VEN_Venta AS v
            INNER JOIN dbo.ADM_Cliente AS c
                ON c.IdCliente = v.IdCliente
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = c.IdPersona
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = v.TipoComprobante
            WHERE v.IdVenta = @IdRegistroComprobante
              AND v.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND v.TipoComprobante <> '07';

            SELECT
                @SaldoNotaCredito = v.Saldo,
                @ImporteTotalNotaCredito = v.ImporteTotal,
                @IdMonedaNotaCredito = v.IdMoneda,
                @TipoComprobanteNc = v.TipoComprobante,
                @SerieNc = v.Serie,
                @NumeroNc = v.Numero,
                @DescripcionTipoNc = tc.Descripcion
            FROM dbo.VEN_Venta AS v
            INNER JOIN dbo.ADM_Cliente AS c
                ON c.IdCliente = v.IdCliente
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = c.IdPersona
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = v.TipoComprobante
            WHERE v.IdVenta = @IdRegistroNotaCredito
              AND v.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND v.TipoComprobante = '07';
        END
        ELSE
        BEGIN
            SELECT
                @SaldoComprobante = c.Saldo,
                @ImporteTotalComprobante = c.ImporteTotal,
                @IdMoneda = c.IdMoneda,
                @TipoCambioDocumento = c.TipoCambio,
                @TipoComprobanteAplicado = c.TipoComprobante,
                @SerieAplicado = c.Serie,
                @NumeroAplicado = c.Numero,
                @NumeroDocumentoPersona = pe.NumeroDocumento,
                @DescripcionTipoAplicado = tc.Descripcion
            FROM dbo.COM_Compra AS c
            INNER JOIN dbo.ADM_Proveedor AS p
                ON p.IdProveedor = c.IdProveedor
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = p.IdPersona
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = c.TipoComprobante
            WHERE c.IdCompra = @IdRegistroComprobante
              AND c.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND c.TipoComprobante <> '07';

            SELECT
                @SaldoNotaCredito = c.Saldo,
                @ImporteTotalNotaCredito = c.ImporteTotal,
                @IdMonedaNotaCredito = c.IdMoneda,
                @TipoComprobanteNc = c.TipoComprobante,
                @SerieNc = c.Serie,
                @NumeroNc = c.Numero,
                @DescripcionTipoNc = tc.Descripcion
            FROM dbo.COM_Compra AS c
            INNER JOIN dbo.ADM_Proveedor AS p
                ON p.IdProveedor = c.IdProveedor
            INNER JOIN dbo.ADM_Persona AS pe
                ON pe.IdPersona = p.IdPersona
            INNER JOIN dbo.ADM_TipoComprobante AS tc
                ON tc.CodigoTipoComprobante = c.TipoComprobante
            WHERE c.IdCompra = @IdRegistroNotaCredito
              AND c.IdEmpresa = @IdEmpresa
              AND pe.IdPersona = @IdPersona
              AND c.TipoComprobante = '07';
        END;

        IF @TipoComprobanteAplicado IS NULL
        BEGIN
            RAISERROR(N'El comprobante pendiente seleccionado no existe, no pertenece a la persona o no es valido para aplicar.', 16, 1);
        END;

        IF @TipoComprobanteNc IS NULL
        BEGIN
            RAISERROR(N'La nota de credito seleccionada no existe, no pertenece a la persona o no es valida para aplicar.', 16, 1);
        END;

        IF @SaldoComprobante <= 0 OR @SaldoNotaCredito <= 0
        BEGIN
            RAISERROR(N'El comprobante o la nota de credito ya no tienen saldo pendiente.', 16, 1);
        END;

        IF @ImporteAplicado > @SaldoComprobante
        BEGIN
            RAISERROR(N'El importe aplicado no puede exceder el saldo pendiente del comprobante.', 16, 1);
        END;

        IF @ImporteAplicado > @SaldoNotaCredito
        BEGIN
            RAISERROR(N'El importe aplicado no puede exceder el saldo pendiente de la nota de credito.', 16, 1);
        END;

        SELECT
            @CodigoMoneda = m.CodigoMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.IdMoneda = @IdMoneda;

        SELECT
            @CodigoMonedaNotaCredito = m.CodigoMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.IdMoneda = @IdMonedaNotaCredito;

        IF NULLIF(LTRIM(RTRIM(@CodigoMoneda)), '') IS NULL
        BEGIN
            RAISERROR(N'No se pudo resolver la moneda del comprobante seleccionado.', 16, 1);
        END;

        IF NULLIF(LTRIM(RTRIM(@CodigoMonedaNotaCredito)), '') IS NULL
        BEGIN
            RAISERROR(N'No se pudo resolver la moneda de la nota de credito seleccionada.', 16, 1);
        END;

        IF UPPER(LTRIM(RTRIM(@CodigoMoneda))) <> UPPER(LTRIM(RTRIM(@CodigoMonedaNotaCredito)))
        BEGIN
            RAISERROR(N'El comprobante y la nota de credito deben estar en la misma moneda para aplicar.', 16, 1);
        END;

        IF @TipoCambioAplicacion > 0
        BEGIN
            SET @TipoCambioAsiento = @TipoCambioAplicacion;
        END
        ELSE IF @TipoCambioDocumento > 0
        BEGIN
            SET @TipoCambioAsiento = @TipoCambioDocumento;
        END;

        SELECT
            @IdOrigen = c.IdOrigen,
            @GeneraAsientoAutomatico = c.GeneraAsientoAutomatico,
            @ConfiguracionActiva = c.Activo
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'APNC'
          AND c.EscenarioOperacion = 'PROVISION';

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'No existe configuracion contable para la provision APNC.', 16, 1);
        END;

        IF @ConfiguracionActiva = 0
        BEGIN
            RAISERROR(N'La configuracion contable APNC esta inactiva.', 16, 1);
        END;

        IF @ModuloOperacion = 'VEN'
        BEGIN
            SELECT
                @IdCuentaComprobante = CASE
                                           WHEN UPPER(LTRIM(RTRIM(@CodigoMoneda))) = 'USD' THEN cfg.IdCuentaVentaDolares
                                           ELSE cfg.IdCuentaVentaSoles
                                       END
            FROM dbo.ADM_TipoComprobante AS tc
            LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
                ON cfg.IdTipoComprobante = tc.IdTipoComprobante
               AND cfg.IdEmpresa = @IdEmpresa
               AND cfg.Activo = 1
            WHERE tc.CodigoTipoComprobante = @TipoComprobanteAplicado;

            SELECT
                @IdCuentaNotaCredito = CASE
                                           WHEN UPPER(LTRIM(RTRIM(@CodigoMoneda))) = 'USD' THEN cfg.IdCuentaVentaDolares
                                           ELSE cfg.IdCuentaVentaSoles
                                       END
            FROM dbo.ADM_TipoComprobante AS tc
            LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
                ON cfg.IdTipoComprobante = tc.IdTipoComprobante
               AND cfg.IdEmpresa = @IdEmpresa
               AND cfg.Activo = 1
            WHERE tc.CodigoTipoComprobante = @TipoComprobanteNc;
        END
        ELSE
        BEGIN
            SELECT
                @IdCuentaComprobante = CASE
                                           WHEN UPPER(LTRIM(RTRIM(@CodigoMoneda))) = 'USD' THEN cfg.IdCuentaCompraDolares
                                           ELSE cfg.IdCuentaCompraSoles
                                       END
            FROM dbo.ADM_TipoComprobante AS tc
            LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
                ON cfg.IdTipoComprobante = tc.IdTipoComprobante
               AND cfg.IdEmpresa = @IdEmpresa
               AND cfg.Activo = 1
            WHERE tc.CodigoTipoComprobante = @TipoComprobanteAplicado;

            SELECT
                @IdCuentaNotaCredito = CASE
                                           WHEN UPPER(LTRIM(RTRIM(@CodigoMoneda))) = 'USD' THEN cfg.IdCuentaCompraDolares
                                           ELSE cfg.IdCuentaCompraSoles
                                       END
            FROM dbo.ADM_TipoComprobante AS tc
            LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
                ON cfg.IdTipoComprobante = tc.IdTipoComprobante
               AND cfg.IdEmpresa = @IdEmpresa
               AND cfg.Activo = 1
            WHERE tc.CodigoTipoComprobante = @TipoComprobanteNc;
        END;

        IF @IdCuentaComprobante IS NULL OR @IdCuentaNotaCredito IS NULL
        BEGIN
            RAISERROR(N'No existe una cuenta contable configurada para el comprobante o la nota de credito en la moneda seleccionada.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS p
            WHERE p.IdPlanCuenta IN (@IdCuentaComprobante, @IdCuentaNotaCredito)
              AND (p.IdEmpresa <> @IdEmpresa OR p.Estado = 0 OR p.AceptaMovimiento = 0)
        )
        BEGIN
            RAISERROR(N'Las cuentas configuradas para la aplicacion no son validas para la empresa activa.', 16, 1);
        END;

        SET @GlosaAsiento = LEFT(CONCAT(N'Aplicacion NC ', @TipoComprobanteAplicado, N' ', @SerieAplicado, N'-', @NumeroAplicado, N' / ', @TipoComprobanteNc, N' ', @SerieNc, N'-', @NumeroNc), 500);
        SET @ReferenciaExterna = LEFT(CONCAT(@TipoComprobanteAplicado, N' ', @SerieAplicado, N'-', @NumeroAplicado, N' / ', @TipoComprobanteNc, N' ', @SerieNc, N'-', @NumeroNc), 100);

        IF @ModuloOperacion = 'VEN'
        BEGIN
            INSERT INTO @DetalleAsiento
            (
                Item,
                IdPlanCuenta,
                GlosaDetalle,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber
            )
            VALUES
            (
                1,
                @IdCuentaNotaCredito,
                LEFT(CONCAT(N'Aplicacion NC ', @TipoComprobanteNc, N' ', @SerieNc, N'-', @NumeroNc), 300),
                @DescripcionTipoNc,
                @NumeroDocumentoPersona,
                @SerieNc,
                LEFT(CONCAT(@SerieNc, N'-', @NumeroNc), 100),
                @TipoCambioAsiento,
                @ImporteAplicado,
                0
            ),
            (
                2,
                @IdCuentaComprobante,
                LEFT(CONCAT(N'Aplicacion comprobante ', @TipoComprobanteAplicado, N' ', @SerieAplicado, N'-', @NumeroAplicado), 300),
                @DescripcionTipoAplicado,
                @NumeroDocumentoPersona,
                @SerieAplicado,
                LEFT(CONCAT(@SerieAplicado, N'-', @NumeroAplicado), 100),
                @TipoCambioAsiento,
                0,
                @ImporteAplicado
            );
        END
        ELSE
        BEGIN
            INSERT INTO @DetalleAsiento
            (
                Item,
                IdPlanCuenta,
                GlosaDetalle,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber
            )
            VALUES
            (
                1,
                @IdCuentaComprobante,
                LEFT(CONCAT(N'Aplicacion comprobante ', @TipoComprobanteAplicado, N' ', @SerieAplicado, N'-', @NumeroAplicado), 300),
                @DescripcionTipoAplicado,
                @NumeroDocumentoPersona,
                @SerieAplicado,
                LEFT(CONCAT(@SerieAplicado, N'-', @NumeroAplicado), 100),
                @TipoCambioAsiento,
                @ImporteAplicado,
                0
            ),
            (
                2,
                @IdCuentaNotaCredito,
                LEFT(CONCAT(N'Aplicacion NC ', @TipoComprobanteNc, N' ', @SerieNc, N'-', @NumeroNc), 300),
                @DescripcionTipoNc,
                @NumeroDocumentoPersona,
                @SerieNc,
                LEFT(CONCAT(@SerieNc, N'-', @NumeroNc), 100),
                @TipoCambioAsiento,
                0,
                @ImporteAplicado
            );
        END;

        SET @DetalleXml =
        (
            SELECT
                d.Item AS [@Item],
                d.IdPlanCuenta AS [@IdPlanCuenta],
                CASE WHEN d.Debe > 0 THEN 'D' ELSE 'H' END AS [@DH],
                d.GlosaDetalle AS [@GlosaDetalle],
                d.TipoDocumento AS [@TipoDocumento],
                d.NumeroDocumento AS [@NumeroDocumento],
                d.Serie AS [@Serie],
                d.ReferenciaLinea AS [@ReferenciaLinea],
                d.TipoCambioLinea AS [@TipoCambioLinea],
                d.Debe AS [@Debe],
                d.Haber AS [@Haber]
            FROM @DetalleAsiento AS d
            ORDER BY d.Item
            FOR XML PATH('Detalle'), ROOT('Detalles'), TYPE
        );

        BEGIN TRANSACTION;

        IF @GeneraAsientoAutomatico = 1
        BEGIN
            BEGIN TRY
                INSERT INTO @ResultadoAsiento
                (
                    IdAsiento,
                    Periodo,
                    NumeroAsiento,
                    TotalDebe,
                    TotalHaber,
                    Estado
                )
                EXEC dbo.usp_CON_GuardarAsientoManual
                    @IdAsiento = NULL,
                    @IdEmpresa = @IdEmpresa,
                    @IdOrigen = @IdOrigen,
                    @FechaEmision = @FechaAplicacion,
                    @FechaAsiento = @FechaAplicacion,
                    @Glosa = @GlosaAsiento,
                    @IdMoneda = @IdMoneda,
                    @TipoCambio = @TipoCambioAsiento,
                    @ReferenciaExterna = @ReferenciaExterna,
                    @Observacion = @Observacion,
                    @DetalleXml = @DetalleXml,
                    @UsuarioRegistro = @UsuarioRegistro;
            END TRY
            BEGIN CATCH
                DECLARE @ErrorAsiento NVARCHAR(4000) = ERROR_MESSAGE();
                RAISERROR(N'No se pudo generar el asiento de la aplicacion. %s', 16, 1, @ErrorAsiento);
            END CATCH;

            SELECT TOP (1)
                @IdAsiento = r.IdAsiento,
                @NumeroAsiento = r.NumeroAsiento
            FROM @ResultadoAsiento AS r;

            IF @IdAsiento IS NOT NULL
            BEGIN
                UPDATE dbo.CON_Asiento
                SET Estado = N'PROVISIONADO'
                WHERE IdAsiento = @IdAsiento
                  AND IdEmpresa = @IdEmpresa;
            END;
        END;

        INSERT INTO dbo.CON_AplicacionNotaCredito
        (
            IdEmpresa,
            ModuloOperacion,
            IdPersona,
            FechaAplicacion,
            IdRegistroComprobante,
            IdRegistroNotaCredito,
            IdMoneda,
            TipoCambio,
            ImporteAplicado,
            IdAsiento,
            Glosa,
            Observacion,
            Activo,
            UsuarioRegistro
        )
        VALUES
        (
            @IdEmpresa,
            @ModuloOperacion,
            @IdPersona,
            @FechaAplicacion,
            @IdRegistroComprobante,
            @IdRegistroNotaCredito,
            @IdMoneda,
            @TipoCambioAsiento,
            @ImporteAplicado,
            @IdAsiento,
            LTRIM(RTRIM(@Glosa)),
            NULLIF(LTRIM(RTRIM(@Observacion)), N''),
            1,
            @UsuarioRegistro
        );

        SET @IdAplicacionNotaCredito = SCOPE_IDENTITY();

        IF @ModuloOperacion = 'VEN'
        BEGIN
            UPDATE dbo.VEN_Venta
            SET Saldo = CASE
                            WHEN Saldo - @ImporteAplicado < 0 THEN 0
                            ELSE Saldo - @ImporteAplicado
                        END
            WHERE IdEmpresa = @IdEmpresa
              AND IdVenta IN (@IdRegistroComprobante, @IdRegistroNotaCredito);
        END
        ELSE
        BEGIN
            UPDATE dbo.COM_Compra
            SET Saldo = CASE
                            WHEN Saldo - @ImporteAplicado < 0 THEN 0
                            ELSE Saldo - @ImporteAplicado
                        END
            WHERE IdEmpresa = @IdEmpresa
              AND IdCompra IN (@IdRegistroComprobante, @IdRegistroNotaCredito);
        END;

        COMMIT TRANSACTION;

        SELECT
            a.IdAplicacionNotaCredito,
            a.IdEmpresa,
            a.ModuloOperacion,
            a.IdPersona,
            a.FechaAplicacion,
            a.IdRegistroComprobante,
            a.IdRegistroNotaCredito,
            a.IdMoneda,
            m.CodigoMoneda,
            a.TipoCambio,
            a.ImporteAplicado,
            a.IdAsiento,
            ca.NumeroAsiento,
            a.Glosa
        FROM dbo.CON_AplicacionNotaCredito AS a
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = a.IdMoneda
        LEFT JOIN dbo.CON_Asiento AS ca
            ON ca.IdAsiento = a.IdAsiento
        WHERE a.IdAplicacionNotaCredito = @IdAplicacionNotaCredito;

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
