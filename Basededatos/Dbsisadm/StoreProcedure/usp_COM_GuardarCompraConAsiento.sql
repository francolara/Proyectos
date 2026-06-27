-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza una provision de compra y genera su asiento automatico segun la configuracion contable.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Calcula subtotal, totales exonerado/inafecto, cuenta contable y afectacion IGV en la provision de compras.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Usa cuenta del detalle, cuenta por documento/impuesto, valida cuentas activas con movimiento, separa RUC/DNI, tipo comprobante, serie y numero en el asiento e informa fecha de emision en la cabecera contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Inicializa y mantiene el saldo del comprobante de compra igual al importe total al registrar o editar la provision.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Conserva la linea original del detalle y agrega cuentas destino y contrapartida segun la configuracion activa por cuenta contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Genera detraccion opcional en compras, descuenta el saldo principal, mantiene el documento hijo SPOT con su asiento en altas y ediciones, toma la cuenta SPOT desde ADM_ParametroEmpresa (CTADETRACCION) y marca la linea SPOT con tipo de documento Otros.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Guarda tipo documento por codigo en compras y detracciones, usa 00 para Otros y calcula importes por moneda en cada linea del asiento.

CREATE OR ALTER PROCEDURE dbo.usp_COM_GuardarCompraConAsiento
    @IdCompra INT = NULL,
    @IdEmpresa INT,
    @IdProveedor INT,
    @IdConfiguracionContabilizacion INT,
    @FechaEmision DATE,
    @FechaContabilizacion DATE,
    @TipoComprobante VARCHAR(3),
    @Serie VARCHAR(10),
    @Numero VARCHAR(20),
    @IdMoneda INT,
    @TipoCambio DECIMAL(18,6),
    @BaseImponible DECIMAL(18,2),
    @TotalExonerado DECIMAL(18,2),
    @TotalInafecto DECIMAL(18,2),
    @Icbper DECIMAL(18,2),
    @Igv DECIMAL(18,2),
    @Isc DECIMAL(18,2),
    @OtrosTributos DECIMAL(18,2),
    @Redondeo DECIMAL(18,2),
    @ImporteTotal DECIMAL(18,2),
    @TieneDetraccion BIT = 0,
    @IdDetraccionSunat INT = NULL,
    @ImporteDetraccion DECIMAL(18,2) = 0,
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCompraTrabajo INT
        DECLARE @IdAsientoTrabajo INT
        DECLARE @IdCompraDetraccionTrabajo INT
        DECLARE @IdAsientoDetraccionTrabajo INT
        DECLARE @IdOrigen INT
        DECLARE @IdOrigenDetraccion INT
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), YEAR(@FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(@FechaContabilizacion)), 2)
        DECLARE @Ejercicio SMALLINT = YEAR(@FechaContabilizacion)
        DECLARE @Mes TINYINT = MONTH(@FechaContabilizacion)
        DECLARE @NumeroAsiento INT
        DECLARE @NumeroAsientoDetraccion INT
        DECLARE @GlosaAsiento NVARCHAR(500)
        DECLARE @GlosaAsientoDetraccion NVARCHAR(500)
        DECLARE @TotalDebe DECIMAL(18,2)
        DECLARE @TotalHaber DECIMAL(18,2)
        DECLARE @TotalDebeDetraccion DECIMAL(18,2)
        DECLARE @TotalHaberDetraccion DECIMAL(18,2)
        DECLARE @EstadoConfiguracion BIT
        DECLARE @GeneraAsientoAutomatico BIT
        DECLARE @EstadoConfiguracionDetraccion BIT
        DECLARE @GeneraAsientoAutomaticoDetraccion BIT
        DECLARE @SubtotalDetalle DECIMAL(18,2)
        DECLARE @TotalExoneradoDetalle DECIMAL(18,2)
        DECLARE @TotalInafectoDetalle DECIMAL(18,2)
        DECLARE @TotalGravadoDetalle DECIMAL(18,2)
        DECLARE @IdTipoComprobanteTrabajo INT
        DECLARE @CodigoMoneda VARCHAR(10)
        DECLARE @IdCuentaDocumento INT
        DECLARE @IdCuentaIgv INT
        DECLARE @IdCuentaIsc INT
        DECLARE @IdCuentaIcbper INT
        DECLARE @IdCuentaOtros INT
        DECLARE @IdCuentaSpot INT
        DECLARE @NumeroDocumentoProveedor VARCHAR(20)
        DECLARE @DescripcionTipoComprobante NVARCHAR(150)
        DECLARE @CodigoDetraccionSunat VARCHAR(3)
        DECLARE @DescripcionDetraccionSunat NVARCHAR(250)
        DECLARE @PorcentajeDetraccion DECIMAL(7,4)
        DECLARE @ImporteDetraccionCalculado DECIMAL(18,2)
        DECLARE @SaldoCompraAnterior DECIMAL(18,2)
        DECLARE @ImporteTotalAnterior DECIMAL(18,2)
        DECLARE @ImporteDetraccionAnterior DECIMAL(18,2)
        DECLARE @SaldoDetraccionAnterior DECIMAL(18,2)

        IF @BaseImponible < 0
           OR @TotalExonerado < 0
           OR @TotalInafecto < 0
           OR @Icbper < 0
           OR @Igv < 0
           OR @Isc < 0
           OR @OtrosTributos < 0
           OR @Redondeo < 0
           OR @ImporteTotal < 0
        BEGIN
            RAISERROR(N'Los montos de la compra no pueden ser negativos.', 16, 1);
        END;

        IF @ImporteTotal <> (@BaseImponible + @Igv)
        BEGIN
            RAISERROR(N'El importe total debe ser igual a la suma del subtotal e IGV.', 16, 1);
        END;

        IF @TieneDetraccion = 0
        BEGIN
            SET @IdDetraccionSunat = NULL;
            SET @ImporteDetraccion = 0;
        END;

        IF @TieneDetraccion = 1 AND ISNULL(@IdDetraccionSunat, 0) <= 0
        BEGIN
            RAISERROR(N'Debe seleccionar el codigo de detraccion SUNAT.', 16, 1);
        END;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la compra.', 16, 1);
        END;

        SELECT
            @NumeroDocumentoProveedor = per.NumeroDocumento
        FROM dbo.ADM_Proveedor AS p
        INNER JOIN dbo.ADM_Persona AS per
            ON per.IdPersona = p.IdPersona
        WHERE p.IdProveedor = @IdProveedor
          AND p.IdEmpresa = @IdEmpresa
          AND p.Estado = 1
          AND per.IdEmpresa = @IdEmpresa
          AND per.Estado = 1;

        IF @NumeroDocumentoProveedor IS NULL
        BEGIN
            RAISERROR(N'El proveedor seleccionado no existe o no pertenece a la empresa.', 16, 1);
        END;

        SELECT
            @IdTipoComprobanteTrabajo = t.IdTipoComprobante,
            @DescripcionTipoComprobante = t.Descripcion
        FROM dbo.ADM_TipoComprobante AS t
        WHERE t.CodigoTipoComprobante = @TipoComprobante
          AND t.UsoCompras = 1
          AND t.Estado = 1;

        IF @IdTipoComprobanteTrabajo IS NULL
        BEGIN
            RAISERROR(N'El tipo de comprobante no existe o no esta habilitado para compras.', 16, 1);
        END;

        SELECT
            @CodigoMoneda = m.CodigoMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.IdMoneda = @IdMoneda
          AND m.Estado = 1;

        IF @CodigoMoneda IS NULL
        BEGIN
            RAISERROR(N'La moneda seleccionada no existe o no esta activa.', 16, 1);
        END;

        IF @TieneDetraccion = 1
        BEGIN
            SELECT
                @CodigoDetraccionSunat = d.CodigoSunat,
                @DescripcionDetraccionSunat = d.Descripcion,
                @PorcentajeDetraccion = d.Porcentaje
            FROM dbo.ADM_DetraccionSunat AS d
            WHERE d.IdDetraccionSunat = @IdDetraccionSunat
              AND d.Estado = 1;

            IF @CodigoDetraccionSunat IS NULL
            BEGIN
                RAISERROR(N'La detraccion seleccionada no existe o no esta activa.', 16, 1);
            END;

            SET @ImporteDetraccionCalculado = ROUND(@ImporteTotal * (@PorcentajeDetraccion / 100.0), 2);

            IF @ImporteDetraccion <> @ImporteDetraccionCalculado
            BEGIN
                RAISERROR(N'El importe de detraccion no coincide con el porcentaje configurado para el codigo SUNAT seleccionado.', 16, 1);
            END;

            IF @ImporteDetraccion <= 0 OR @ImporteDetraccion >= @ImporteTotal
            BEGIN
                RAISERROR(N'La detraccion debe ser mayor a cero y menor al importe total de la compra.', 16, 1);
            END;
        END;

        SELECT
            @IdOrigen = c.IdOrigen,
            @EstadoConfiguracion = c.Activo,
            @GeneraAsientoAutomatico = c.GeneraAsientoAutomatico
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
          AND c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'COM';

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'La configuracion contable indicada no existe para compras en la empresa activa.', 16, 1);
        END;

        IF @EstadoConfiguracion = 0
        BEGIN
            RAISERROR(N'La configuracion contable seleccionada esta inactiva.', 16, 1);
        END;

        IF @GeneraAsientoAutomatico = 0
        BEGIN
            RAISERROR(N'La configuracion seleccionada no esta habilitada para generar asiento automatico.', 16, 1);
        END;

        IF @TieneDetraccion = 1
        BEGIN
            SELECT
                @IdOrigenDetraccion = c.IdOrigen,
                @EstadoConfiguracionDetraccion = c.Activo,
                @GeneraAsientoAutomaticoDetraccion = c.GeneraAsientoAutomatico
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.ModuloOperacion = 'DET'
              AND c.EscenarioOperacion = 'PROVISION';

            IF @IdOrigenDetraccion IS NULL
            BEGIN
                RAISERROR(N'No existe configuracion contable activa para detracciones en la empresa.', 16, 1);
            END;

            IF @EstadoConfiguracionDetraccion = 0
            BEGIN
                RAISERROR(N'La configuracion contable de detracciones esta inactiva.', 16, 1);
            END;

            IF @GeneraAsientoAutomaticoDetraccion = 0
            BEGIN
                RAISERROR(N'La configuracion contable de detracciones no esta habilitada para generar asiento automatico.', 16, 1);
            END;
        END;

        DECLARE @DetalleCompra TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            IdTipoAfectacionIGV INT NOT NULL,
            Descripcion NVARCHAR(250) NOT NULL,
            Cantidad DECIMAL(18,4) NOT NULL,
            ValorUnitario DECIMAL(18,6) NOT NULL,
            ImporteBruto DECIMAL(18,2) NOT NULL
        );

        INSERT INTO @DetalleCompra
        (
            Item,
            IdPlanCuenta,
            IdTipoAfectacionIGV,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto
        )
        SELECT
            T.N.value('@Item', 'smallint'),
            T.N.value('@IdPlanCuenta', 'int'),
            T.N.value('@IdTipoAfectacionIGV', 'int'),
            T.N.value('@Descripcion', 'nvarchar(250)'),
            T.N.value('@Cantidad', 'decimal(18,4)'),
            T.N.value('@ValorUnitario', 'decimal(18,6)'),
            T.N.value('@ImporteBruto', 'decimal(18,2)')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @DetalleCompra
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos una linea en la compra.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @DetalleCompra AS d
            WHERE d.Item < 1
               OR d.Cantidad <= 0
               OR d.ValorUnitario < 0
               OR d.ImporteBruto < 0
               OR d.IdPlanCuenta <= 0
               OR d.IdTipoAfectacionIGV <= 0
        )
        BEGIN
            RAISERROR(N'El detalle de la compra contiene valores no validos.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @DetalleCompra AS d
            LEFT JOIN dbo.CON_PlanCuenta AS p
                ON p.IdPlanCuenta = d.IdPlanCuenta
               AND p.IdEmpresa = @IdEmpresa
               AND p.AceptaMovimiento = 1
               AND p.Estado = 1
            WHERE p.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR(N'La cuenta contable seleccionada en el detalle de la compra no es valida. Verifique que pertenezca a la empresa, este activa y acepte movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @DetalleCompra AS d
            LEFT JOIN dbo.CON_TipoAfectacionIGV AS a
                ON a.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV
               AND a.Estado = 1
            WHERE a.IdTipoAfectacionIGV IS NULL
        )
        BEGIN
            RAISERROR(N'El detalle contiene un tipo de afectacion IGV invalido.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.Item
            FROM @DetalleCompra AS d
            GROUP BY
                d.Item
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten items duplicados en el detalle de la compra.', 16, 1);
        END;

        SELECT
            @SubtotalDetalle = ISNULL(SUM(d.ImporteBruto), 0),
            @TotalExoneradoDetalle = ISNULL(SUM(CASE WHEN a.CodigoSunat LIKE '2%' THEN d.ImporteBruto ELSE 0 END), 0),
            @TotalInafectoDetalle = ISNULL(SUM(CASE WHEN a.CodigoSunat LIKE '3%' THEN d.ImporteBruto ELSE 0 END), 0),
            @TotalGravadoDetalle = ISNULL(SUM(CASE WHEN a.CodigoSunat LIKE '1%' THEN d.ImporteBruto ELSE 0 END), 0)
        FROM @DetalleCompra AS d
        INNER JOIN dbo.CON_TipoAfectacionIGV AS a
            ON a.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV;

        IF @BaseImponible <> @SubtotalDetalle
        BEGIN
            RAISERROR(N'El subtotal debe coincidir con la suma del detalle.', 16, 1);
        END;

        IF @TotalExonerado <> @TotalExoneradoDetalle
        BEGIN
            RAISERROR(N'El total exonerado debe coincidir con la afectacion IGV del detalle.', 16, 1);
        END;

        IF @TotalInafecto <> @TotalInafectoDetalle
        BEGIN
            RAISERROR(N'El total inafecto debe coincidir con la afectacion IGV del detalle.', 16, 1);
        END;

        IF @Igv <> ROUND(@TotalGravadoDetalle * 0.18, 2)
        BEGIN
            RAISERROR(N'El IGV debe calcularse con base en los items gravados del detalle.', 16, 1);
        END;

        SELECT
            @IdCuentaDocumento = CASE
                WHEN @CodigoMoneda = 'USD' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdCuentaCompraDolares END, t.IdCuentaCompraDolares)
                ELSE COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdCuentaCompraSoles END, t.IdCuentaCompraSoles)
            END
        FROM dbo.ADM_TipoComprobante AS t
        LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoComprobante = t.IdTipoComprobante
           AND cfg.IdEmpresa = @IdEmpresa
        WHERE t.IdTipoComprobante = @IdTipoComprobanteTrabajo;

        IF @IdCuentaDocumento IS NULL
        BEGIN
            RAISERROR(N'No existe una cuenta contable configurada para compras en el tipo de comprobante y moneda seleccionados.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS p
            WHERE p.IdPlanCuenta = @IdCuentaDocumento
              AND p.IdEmpresa = @IdEmpresa
              AND p.Estado = 1
              AND p.AceptaMovimiento = 1
        )
        BEGIN
            RAISERROR(N'La cuenta contable configurada para el documento de compra no pertenece a la empresa, no esta activa o no acepta movimiento.', 16, 1);
        END;

        SELECT
            @IdCuentaIgv = MAX(CASE WHEN i.CodigoSunat = 'IGV' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdPlanCuenta END, i.IdPlanCuenta) END),
            @IdCuentaIsc = MAX(CASE WHEN i.CodigoSunat = 'ISC' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdPlanCuenta END, i.IdPlanCuenta) END),
            @IdCuentaIcbper = MAX(CASE WHEN i.CodigoSunat = 'ICBPER' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdPlanCuenta END, i.IdPlanCuenta) END),
            @IdCuentaOtros = MAX(CASE WHEN i.CodigoSunat = 'OTROS' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdPlanCuenta END, i.IdPlanCuenta) END)
        FROM dbo.CON_TipoImpuesto AS i
        LEFT JOIN dbo.CON_TipoImpuestoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoImpuesto = i.IdTipoImpuesto
           AND cfg.IdEmpresa = @IdEmpresa
        WHERE i.Estado = 1;

        SELECT TOP (1)
            @IdCuentaSpot = pc.IdPlanCuenta
        FROM dbo.ADM_ParametroEmpresa AS pe
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.CodigoCuenta = pe.ValorParametro
           AND pc.IdEmpresa = pe.IdEmpresa
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.CodigoParametro = 'CTADETRACCION'
          AND pe.Activo = 1
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1
        ORDER BY
            pe.IdParametroEmpresa DESC;

        IF @Igv > 0
        AND (
            @IdCuentaIgv IS NULL
            OR NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS p
                WHERE p.IdPlanCuenta = @IdCuentaIgv
                  AND p.IdEmpresa = @IdEmpresa
                  AND p.Estado = 1
                  AND p.AceptaMovimiento = 1
            )
        )
        BEGIN
            RAISERROR(N'No existe una cuenta contable valida configurada para IGV en la empresa.', 16, 1);
        END;

        IF @Isc > 0
        AND (
            @IdCuentaIsc IS NULL
            OR NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS p
                WHERE p.IdPlanCuenta = @IdCuentaIsc
                  AND p.IdEmpresa = @IdEmpresa
                  AND p.Estado = 1
                  AND p.AceptaMovimiento = 1
            )
        )
        BEGIN
            RAISERROR(N'No existe una cuenta contable valida configurada para ISC en la empresa.', 16, 1);
        END;

        IF @Icbper > 0
        AND (
            @IdCuentaIcbper IS NULL
            OR NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS p
                WHERE p.IdPlanCuenta = @IdCuentaIcbper
                  AND p.IdEmpresa = @IdEmpresa
                  AND p.Estado = 1
                  AND p.AceptaMovimiento = 1
            )
        )
        BEGIN
            RAISERROR(N'No existe una cuenta contable valida configurada para ICBPER en la empresa.', 16, 1);
        END;

        IF @OtrosTributos > 0
        AND (
            @IdCuentaOtros IS NULL
            OR NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS p
                WHERE p.IdPlanCuenta = @IdCuentaOtros
                  AND p.IdEmpresa = @IdEmpresa
                  AND p.Estado = 1
                  AND p.AceptaMovimiento = 1
            )
        )
        BEGIN
            RAISERROR(N'No existe una cuenta contable valida configurada para otros tributos en la empresa.', 16, 1);
        END;

        IF @TieneDetraccion = 1
        AND (
            @IdCuentaSpot IS NULL
            OR NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS p
                WHERE p.IdPlanCuenta = @IdCuentaSpot
                  AND p.IdEmpresa = @IdEmpresa
                  AND p.Estado = 1
                  AND p.AceptaMovimiento = 1
            )
        )
        BEGIN
            RAISERROR(N'No existe una cuenta contable valida configurada para el parametro CTADETRACCION en la empresa.', 16, 1);
        END;

        DECLARE @AsientoDetalle TABLE
        (
            Item SMALLINT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL
        );

        DECLARE @AsientoDetraccionDetalle TABLE
        (
            Item SMALLINT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL
        );

        DECLARE @DetalleDestinoBase TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            ImporteBase DECIMAL(18,2) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL
        );

        DECLARE @CuentaDestinoDetalle TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            Orden SMALLINT NOT NULL,
            IdPlanCuentaDestinoCargo INT NOT NULL,
            IdPlanCuentaDestinoAbono INT NOT NULL,
            Porcentaje DECIMAL(7,4) NOT NULL,
            EsUltimo BIT NOT NULL
        );

        INSERT INTO @AsientoDetalle
        (
            IdPlanCuenta,
            Debe,
            Haber,
            GlosaDetalle
        )
        SELECT
            d.IdPlanCuenta,
            SUM(d.ImporteBruto) AS Debe,
            0 AS Haber,
            CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Detalle')
        FROM @DetalleCompra AS d
        GROUP BY d.IdPlanCuenta;

        INSERT INTO @DetalleDestinoBase
        (
            IdPlanCuentaOrigen,
            ImporteBase,
            GlosaDetalle
        )
        SELECT
            d.IdPlanCuenta,
            SUM(d.ImporteBruto),
            CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Detalle')
        FROM @DetalleCompra AS d
        GROUP BY d.IdPlanCuenta;

        IF @Igv > 0
        BEGIN
            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle
            )
            VALUES
            (
                @IdCuentaIgv,
                @Igv,
                0,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / IGV')
            );
        END;

        IF @Isc > 0
        BEGIN
            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle
            )
            VALUES
            (
                @IdCuentaIsc,
                @Isc,
                0,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ISC')
            );
        END;

        IF @Icbper > 0
        BEGIN
            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle
            )
            VALUES
            (
                @IdCuentaIcbper,
                @Icbper,
                0,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ICBPER')
            );
        END;

        IF @OtrosTributos > 0
        BEGIN
            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle
            )
            VALUES
            (
                @IdCuentaOtros,
                @OtrosTributos,
                0,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Otros tributos')
            );
        END;

        INSERT INTO @AsientoDetalle
        (
            IdPlanCuenta,
            Debe,
            Haber,
            GlosaDetalle
        )
        VALUES
        (
            @IdCuentaDocumento,
            0,
            @ImporteTotal,
            CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Documento')
        );

        IF @TieneDetraccion = 1
        BEGIN
            INSERT INTO @AsientoDetraccionDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea
            )
            VALUES
            (
                @IdCuentaDocumento,
                @ImporteDetraccion,
                0,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Aplicacion detraccion'),
                @TipoComprobante,
                @NumeroDocumentoProveedor,
                @Serie,
                @Numero
            );

            INSERT INTO @AsientoDetraccionDetalle
            (
                IdPlanCuenta,
                Debe,
                Haber,
                GlosaDetalle,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea
            )
            VALUES
            (
                @IdCuentaSpot,
                0,
                @ImporteDetraccion,
                CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Detraccion'),
                N'00',
                @NumeroDocumentoProveedor,
                @Serie,
                @Numero
            );
        END;

        INSERT INTO @CuentaDestinoDetalle
        (
            IdPlanCuentaOrigen,
            Orden,
            IdPlanCuentaDestinoCargo,
            IdPlanCuentaDestinoAbono,
            Porcentaje,
            EsUltimo
        )
        SELECT
            r.IdPlanCuentaOrigen,
            d.Orden,
            d.IdPlanCuentaDestinoCargo,
            d.IdPlanCuentaDestinoAbono,
            d.Porcentaje,
            CASE
                WHEN ROW_NUMBER() OVER (PARTITION BY r.IdPlanCuentaOrigen ORDER BY d.Orden DESC) = 1 THEN 1
                ELSE 0
            END
        FROM dbo.CON_CuentaDestinoRegla AS r
        INNER JOIN dbo.CON_CuentaDestinoReglaDetalle AS d
            ON d.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
           AND d.Activo = 1
        INNER JOIN @DetalleDestinoBase AS b
            ON b.IdPlanCuentaOrigen = r.IdPlanCuentaOrigen
        WHERE r.IdEmpresa = @IdEmpresa
          AND r.Activo = 1;

        IF EXISTS
        (
            SELECT 1
            FROM @CuentaDestinoDetalle AS d
            LEFT JOIN dbo.CON_PlanCuenta AS cargo
                ON cargo.IdPlanCuenta = d.IdPlanCuentaDestinoCargo
               AND cargo.IdEmpresa = @IdEmpresa
               AND cargo.Estado = 1
               AND cargo.AceptaMovimiento = 1
            LEFT JOIN dbo.CON_PlanCuenta AS abono
                ON abono.IdPlanCuenta = d.IdPlanCuentaDestinoAbono
               AND abono.IdEmpresa = @IdEmpresa
               AND abono.Estado = 1
               AND abono.AceptaMovimiento = 1
            WHERE cargo.IdPlanCuenta IS NULL
               OR abono.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR(N'Existe una configuracion activa de cuentas destino con cuentas cargo o abono invalidas para la empresa.', 16, 1);
        END;

        DECLARE @IdPlanCuentaOrigenDestino INT
        DECLARE @ImporteBaseDestino DECIMAL(18,2)
        DECLARE @GlosaBaseDestino NVARCHAR(300)
        DECLARE @IdCuentaCargoDestino INT
        DECLARE @IdCuentaAbonoDestino INT
        DECLARE @PorcentajeDestino DECIMAL(7,4)
        DECLARE @EsUltimoDestino BIT
        DECLARE @ImporteDistribuidoDestino DECIMAL(18,2)
        DECLARE @ImporteTramoDestino DECIMAL(18,2)

        DECLARE cursor_linea_destino CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            b.IdPlanCuentaOrigen,
            b.ImporteBase,
            b.GlosaDetalle
        FROM @DetalleDestinoBase AS b
        WHERE b.ImporteBase > 0
          AND EXISTS
          (
              SELECT 1
              FROM @CuentaDestinoDetalle AS d
              WHERE d.IdPlanCuentaOrigen = b.IdPlanCuentaOrigen
          );

        OPEN cursor_linea_destino;

        FETCH NEXT FROM cursor_linea_destino
        INTO @IdPlanCuentaOrigenDestino, @ImporteBaseDestino, @GlosaBaseDestino;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @ImporteDistribuidoDestino = 0;

            DECLARE cursor_tramo_destino CURSOR LOCAL FAST_FORWARD FOR
            SELECT
                d.IdPlanCuentaDestinoCargo,
                d.IdPlanCuentaDestinoAbono,
                d.Porcentaje,
                d.EsUltimo
            FROM @CuentaDestinoDetalle AS d
            WHERE d.IdPlanCuentaOrigen = @IdPlanCuentaOrigenDestino
            ORDER BY
                d.Orden ASC;

            OPEN cursor_tramo_destino;

            FETCH NEXT FROM cursor_tramo_destino
            INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;

            WHILE @@FETCH_STATUS = 0
            BEGIN
                SET @ImporteTramoDestino = CASE
                                               WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDestino - @ImporteDistribuidoDestino
                                               ELSE ROUND(@ImporteBaseDestino * (@PorcentajeDestino / 100.0), 2)
                                           END;

                IF @ImporteTramoDestino <> 0
                BEGIN
                    INSERT INTO @AsientoDetalle
                    (
                        IdPlanCuenta,
                        Debe,
                        Haber,
                        GlosaDetalle
                    )
                    VALUES
                    (
                        @IdCuentaCargoDestino,
                        @ImporteTramoDestino,
                        0,
                        LEFT(CONCAT(ISNULL(@GlosaBaseDestino, N'Distribucion'), N' / Destino'), 300)
                    );

                    INSERT INTO @AsientoDetalle
                    (
                        IdPlanCuenta,
                        Debe,
                        Haber,
                        GlosaDetalle
                    )
                    VALUES
                    (
                        @IdCuentaAbonoDestino,
                        0,
                        @ImporteTramoDestino,
                        LEFT(CONCAT(ISNULL(@GlosaBaseDestino, N'Distribucion'), N' / Contrapartida'), 300)
                    );
                END;

                SET @ImporteDistribuidoDestino = @ImporteDistribuidoDestino + @ImporteTramoDestino;

                FETCH NEXT FROM cursor_tramo_destino
                INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
            END;

            CLOSE cursor_tramo_destino;
            DEALLOCATE cursor_tramo_destino;

            FETCH NEXT FROM cursor_linea_destino
            INTO @IdPlanCuentaOrigenDestino, @ImporteBaseDestino, @GlosaBaseDestino;
        END;

        CLOSE cursor_linea_destino;
        DEALLOCATE cursor_linea_destino;

        DELETE FROM @AsientoDetalle
        WHERE Debe = 0
          AND Haber = 0;

        IF NOT EXISTS
        (
            SELECT 1
            FROM @AsientoDetalle
        )
        BEGIN
            RAISERROR(N'No se pudieron generar lineas contables para la compra con la configuracion de documento, impuestos y detalle.', 16, 1);
        END;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @AsientoDetalle AS d;

        IF @TotalDebe <> @TotalHaber
        BEGIN
            RAISERROR(N'La configuracion contable de compras no genera un asiento cuadrado para los importes ingresados.', 16, 1);
        END;

        SET @GlosaAsiento = CONCAT(N'Compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero);
        SET @GlosaAsientoDetraccion = CONCAT(N'Detraccion compra ', @TipoComprobante, N' ', @Serie, N'-', @Numero);

        IF @TieneDetraccion = 1
        BEGIN
            SELECT
                @TotalDebeDetraccion = SUM(d.Debe),
                @TotalHaberDetraccion = SUM(d.Haber)
            FROM @AsientoDetraccionDetalle AS d;

            IF ISNULL(@TotalDebeDetraccion, 0) <> ISNULL(@TotalHaberDetraccion, 0)
            BEGIN
                RAISERROR(N'La configuracion contable de detracciones no genera un asiento cuadrado para la compra.', 16, 1);
            END;
        END;

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

        BEGIN TRAN;

        IF @IdCompra IS NULL
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @Periodo
            )
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = UltimoNumero + 1,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @Periodo;

                SELECT
                    @NumeroAsiento = c.UltimoNumero
                FROM dbo.CON_CorrelativoAsiento AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @Periodo;
            END
            ELSE
            BEGIN
                INSERT INTO dbo.CON_CorrelativoAsiento
                (
                    IdEmpresa,
                    IdOrigen,
                    Periodo,
                    UltimoNumero,
                    FechaActualizacion,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdOrigen,
                    @Periodo,
                    1,
                    SYSDATETIME(),
                    @UsuarioRegistro
                );

                SET @NumeroAsiento = 1;
            END;

            INSERT INTO dbo.CON_Asiento
            (
                IdEmpresa,
                IdOrigen,
                Ejercicio,
                Mes,
                Periodo,
                NumeroAsiento,
                FechaEmision,
                FechaAsiento,
                Glosa,
                IdMoneda,
                TipoCambio,
                TotalDebe,
                TotalHaber,
                Estado,
                ReferenciaExterna,
                Observacion,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdOrigen,
                @Ejercicio,
                @Mes,
                @Periodo,
                @NumeroAsiento,
                @FechaEmision,
                @FechaContabilizacion,
                @GlosaAsiento,
                @IdMoneda,
                @TipoCambio,
                @TotalDebe,
                @TotalHaber,
                N'PROVISIONADO',
                CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdAsientoTrabajo = SCOPE_IDENTITY();

            INSERT INTO dbo.COM_Compra
            (
                IdEmpresa,
                IdProveedor,
                IdConfiguracionContabilizacion,
                IdAsiento,
                FechaEmision,
                FechaContabilizacion,
                TipoComprobante,
                Serie,
                Numero,
                IdMoneda,
                TipoCambio,
                BaseImponible,
                TotalExonerado,
                TotalInafecto,
                Icbper,
                Igv,
                Isc,
                OtrosTributos,
                Redondeo,
                ImporteTotal,
                Saldo,
                TieneDetraccion,
                IdDetraccionSunat,
                PorcentajeDetraccion,
                ImporteDetraccion,
                Observacion,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdProveedor,
                @IdConfiguracionContabilizacion,
                @IdAsientoTrabajo,
                @FechaEmision,
                @FechaContabilizacion,
                @TipoComprobante,
                @Serie,
                @Numero,
                @IdMoneda,
                @TipoCambio,
                @BaseImponible,
                @TotalExonerado,
                @TotalInafecto,
                @Icbper,
                @Igv,
                @Isc,
                @OtrosTributos,
                @Redondeo,
                @ImporteTotal,
                @ImporteTotal - @ImporteDetraccion,
                @TieneDetraccion,
                @IdDetraccionSunat,
                CASE WHEN @TieneDetraccion = 1 THEN @PorcentajeDetraccion ELSE 0 END,
                CASE WHEN @TieneDetraccion = 1 THEN @ImporteDetraccion ELSE 0 END,
                @Observacion,
                N'PROVISIONADO',
                @UsuarioRegistro
            );

            SET @IdCompraTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT
                @IdCompraTrabajo = c.IdCompra,
                @IdAsientoTrabajo = c.IdAsiento,
                @ImporteTotalAnterior = c.ImporteTotal,
                @SaldoCompraAnterior = c.Saldo,
                @ImporteDetraccionAnterior = c.ImporteDetraccion,
                @IdCompraDetraccionTrabajo = cd.IdCompraDetraccion,
                @IdAsientoDetraccionTrabajo = cd.IdAsiento,
                @SaldoDetraccionAnterior = cd.Saldo
            FROM dbo.COM_Compra AS c
            LEFT JOIN dbo.COM_CompraDetraccion AS cd
                ON cd.IdCompra = c.IdCompra
            WHERE c.IdCompra = @IdCompra
              AND c.IdEmpresa = @IdEmpresa;

            IF @IdCompraTrabajo IS NULL
            BEGIN
                RAISERROR(N'La compra indicada no existe para la empresa activa.', 16, 1);
            END;

            IF ISNULL(@SaldoCompraAnterior, 0) < ISNULL(@ImporteTotalAnterior, 0) - ISNULL(@ImporteDetraccionAnterior, 0)
            BEGIN
                RAISERROR(N'La compra ya tiene pagos aplicados y no puede modificarse desde provisiones.', 16, 1);
            END;

            IF @IdCompraDetraccionTrabajo IS NOT NULL
               AND ISNULL(@SaldoDetraccionAnterior, 0) < ISNULL(@ImporteDetraccionAnterior, 0)
            BEGIN
                RAISERROR(N'La detraccion vinculada ya tiene pagos aplicados y no puede modificarse desde provisiones.', 16, 1);
            END;

            UPDATE dbo.COM_Compra
            SET IdProveedor = @IdProveedor,
                IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion,
                FechaEmision = @FechaEmision,
                FechaContabilizacion = @FechaContabilizacion,
                TipoComprobante = @TipoComprobante,
                Serie = @Serie,
                Numero = @Numero,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                BaseImponible = @BaseImponible,
                TotalExonerado = @TotalExonerado,
                TotalInafecto = @TotalInafecto,
                Icbper = @Icbper,
                Igv = @Igv,
                Isc = @Isc,
                OtrosTributos = @OtrosTributos,
                Redondeo = @Redondeo,
                ImporteTotal = @ImporteTotal,
                Saldo = @ImporteTotal - @ImporteDetraccion,
                TieneDetraccion = @TieneDetraccion,
                IdDetraccionSunat = @IdDetraccionSunat,
                PorcentajeDetraccion = CASE WHEN @TieneDetraccion = 1 THEN @PorcentajeDetraccion ELSE 0 END,
                ImporteDetraccion = CASE WHEN @TieneDetraccion = 1 THEN @ImporteDetraccion ELSE 0 END,
                Observacion = @Observacion,
                Estado = N'PROVISIONADO',
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdCompra = @IdCompraTrabajo;

            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.CON_Asiento
            SET FechaEmision = @FechaEmision,
                FechaAsiento = @FechaContabilizacion,
                Glosa = @GlosaAsiento,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebe,
                TotalHaber = @TotalHaber,
                Estado = N'PROVISIONADO',
                ReferenciaExterna = CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdAsiento = @IdAsientoTrabajo;

            DELETE FROM dbo.COM_CompraDetalle
            WHERE IdCompra = @IdCompraTrabajo;
        END;

        IF @TieneDetraccion = 1
        BEGIN
            IF @IdCompraDetraccionTrabajo IS NULL
            BEGIN
                IF EXISTS
                (
                    SELECT 1
                    FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
                    WHERE c.IdEmpresa = @IdEmpresa
                      AND c.IdOrigen = @IdOrigenDetraccion
                      AND c.Periodo = @Periodo
                )
                BEGIN
                    UPDATE dbo.CON_CorrelativoAsiento
                    SET UltimoNumero = UltimoNumero + 1,
                        FechaActualizacion = SYSDATETIME(),
                        UsuarioRegistro = @UsuarioRegistro
                    WHERE IdEmpresa = @IdEmpresa
                      AND IdOrigen = @IdOrigenDetraccion
                      AND Periodo = @Periodo;

                    SELECT
                        @NumeroAsientoDetraccion = c.UltimoNumero
                    FROM dbo.CON_CorrelativoAsiento AS c
                    WHERE c.IdEmpresa = @IdEmpresa
                      AND c.IdOrigen = @IdOrigenDetraccion
                      AND c.Periodo = @Periodo;
                END
                ELSE
                BEGIN
                    INSERT INTO dbo.CON_CorrelativoAsiento
                    (
                        IdEmpresa,
                        IdOrigen,
                        Periodo,
                        UltimoNumero,
                        FechaActualizacion,
                        UsuarioRegistro
                    )
                    VALUES
                    (
                        @IdEmpresa,
                        @IdOrigenDetraccion,
                        @Periodo,
                        1,
                        SYSDATETIME(),
                        @UsuarioRegistro
                    );

                    SET @NumeroAsientoDetraccion = 1;
                END;

                INSERT INTO dbo.CON_Asiento
                (
                    IdEmpresa,
                    IdOrigen,
                    Ejercicio,
                    Mes,
                    Periodo,
                    NumeroAsiento,
                    FechaEmision,
                    FechaAsiento,
                    Glosa,
                    IdMoneda,
                    TipoCambio,
                    TotalDebe,
                    TotalHaber,
                    Estado,
                    ReferenciaExterna,
                    Observacion,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdOrigenDetraccion,
                    @Ejercicio,
                    @Mes,
                    @Periodo,
                    @NumeroAsientoDetraccion,
                    @FechaEmision,
                    @FechaContabilizacion,
                    @GlosaAsientoDetraccion,
                    @IdMoneda,
                    @TipoCambio,
                    @TotalDebeDetraccion,
                    @TotalHaberDetraccion,
                    N'PROVISIONADO',
                    CONCAT(N'DET ', @TipoComprobante, N' ', @Serie, N'-', @Numero),
                    @Observacion,
                    @UsuarioRegistro
                );

                SET @IdAsientoDetraccionTrabajo = SCOPE_IDENTITY();

                INSERT INTO dbo.COM_CompraDetraccion
                (
                    IdEmpresa,
                    IdCompra,
                    IdProveedor,
                    IdDetraccionSunat,
                    IdAsiento,
                    FechaEmision,
                    FechaContabilizacion,
                    IdMoneda,
                    TipoCambio,
                    CodigoDetraccionSunat,
                    DescripcionDetraccion,
                    PorcentajeDetraccion,
                    ImporteDetraccion,
                    Saldo,
                    ReferenciaDocumento,
                    Observacion,
                    Estado,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdCompraTrabajo,
                    @IdProveedor,
                    @IdDetraccionSunat,
                    @IdAsientoDetraccionTrabajo,
                    @FechaEmision,
                    @FechaContabilizacion,
                    @IdMoneda,
                    @TipoCambio,
                    @CodigoDetraccionSunat,
                    @DescripcionDetraccionSunat,
                    @PorcentajeDetraccion,
                    @ImporteDetraccion,
                    @ImporteDetraccion,
                    CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                    @Observacion,
                    N'PROVISIONADO',
                    @UsuarioRegistro
                );

                SET @IdCompraDetraccionTrabajo = SCOPE_IDENTITY();
            END
            ELSE
            BEGIN
                DELETE FROM dbo.CON_AsientoDetalle
                WHERE IdAsiento = @IdAsientoDetraccionTrabajo;

                UPDATE dbo.CON_Asiento
                SET FechaEmision = @FechaEmision,
                    FechaAsiento = @FechaContabilizacion,
                    Glosa = @GlosaAsientoDetraccion,
                    IdMoneda = @IdMoneda,
                    TipoCambio = @TipoCambio,
                    TotalDebe = @TotalDebeDetraccion,
                    TotalHaber = @TotalHaberDetraccion,
                    Estado = N'PROVISIONADO',
                    ReferenciaExterna = CONCAT(N'DET ', @TipoComprobante, N' ', @Serie, N'-', @Numero),
                    Observacion = @Observacion,
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdAsiento = @IdAsientoDetraccionTrabajo;

                UPDATE dbo.COM_CompraDetraccion
                SET IdProveedor = @IdProveedor,
                    IdDetraccionSunat = @IdDetraccionSunat,
                    IdAsiento = @IdAsientoDetraccionTrabajo,
                    FechaEmision = @FechaEmision,
                    FechaContabilizacion = @FechaContabilizacion,
                    IdMoneda = @IdMoneda,
                    TipoCambio = @TipoCambio,
                    CodigoDetraccionSunat = @CodigoDetraccionSunat,
                    DescripcionDetraccion = @DescripcionDetraccionSunat,
                    PorcentajeDetraccion = @PorcentajeDetraccion,
                    ImporteDetraccion = @ImporteDetraccion,
                    Saldo = @ImporteDetraccion,
                    ReferenciaDocumento = CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                    Observacion = @Observacion,
                    Estado = N'PROVISIONADO',
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdCompraDetraccion = @IdCompraDetraccionTrabajo;
            END;
        END
        ELSE IF @IdCompraDetraccionTrabajo IS NOT NULL
        BEGIN
            DELETE FROM dbo.COM_CompraDetraccion
            WHERE IdCompraDetraccion = @IdCompraDetraccionTrabajo;

            IF @IdAsientoDetraccionTrabajo IS NOT NULL
            BEGIN
                DELETE FROM dbo.CON_AsientoDetalle
                WHERE IdAsiento = @IdAsientoDetraccionTrabajo;

                DELETE FROM dbo.CON_Asiento
                WHERE IdAsiento = @IdAsientoDetraccionTrabajo
                  AND IdEmpresa = @IdEmpresa;
            END;

            SET @IdCompraDetraccionTrabajo = NULL;
            SET @IdAsientoDetraccionTrabajo = NULL;
        END;

        INSERT INTO dbo.CON_AsientoDetalle
        (
            IdAsiento,
            Item,
            IdPlanCuenta,
            GlosaDetalle,
            TipoDocumento,
            NumeroDocumento,
            Serie,
            IdProveedor,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            ReferenciaLinea,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.GlosaDetalle,
            @TipoComprobante,
            @NumeroDocumentoProveedor,
            @Serie,
            @IdProveedor,
            d.Debe,
            d.Haber,
            CASE
                WHEN @CodigoMoneda = 'USD' THEN ROUND(calc.ImporteLinea * calc.TipoCambioAplicado, 2)
                ELSE calc.ImporteLinea
            END,
            CASE
                WHEN @CodigoMoneda = 'USD' THEN calc.ImporteLinea
                ELSE ROUND(calc.ImporteLinea / NULLIF(calc.TipoCambioAplicado, 0), 2)
            END,
            @Numero,
            @UsuarioRegistro
        FROM @AsientoDetalle AS d
        CROSS APPLY
        (
            SELECT
                CASE
                    WHEN d.Debe > 0 THEN d.Debe
                    ELSE d.Haber
                END AS ImporteLinea,
                CASE WHEN @TipoCambio > 0 THEN @TipoCambio ELSE 1 END AS TipoCambioAplicado
        ) AS calc
        ORDER BY
            d.Item ASC;

        IF @TieneDetraccion = 1 AND @IdAsientoDetraccionTrabajo IS NOT NULL
        BEGIN
            INSERT INTO dbo.CON_AsientoDetalle
            (
                IdAsiento,
                Item,
                IdPlanCuenta,
                GlosaDetalle,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                IdProveedor,
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD,
                ReferenciaLinea,
                UsuarioRegistro
            )
            SELECT
                @IdAsientoDetraccionTrabajo,
                d.Item,
                d.IdPlanCuenta,
                d.GlosaDetalle,
                d.TipoDocumento,
                d.NumeroDocumento,
                d.Serie,
                @IdProveedor,
                d.Debe,
                d.Haber,
                CASE
                    WHEN @CodigoMoneda = 'USD' THEN ROUND(calc.ImporteLinea * calc.TipoCambioAplicado, 2)
                    ELSE calc.ImporteLinea
                END,
                CASE
                    WHEN @CodigoMoneda = 'USD' THEN calc.ImporteLinea
                    ELSE ROUND(calc.ImporteLinea / NULLIF(calc.TipoCambioAplicado, 0), 2)
                END,
                d.ReferenciaLinea,
                @UsuarioRegistro
            FROM @AsientoDetraccionDetalle AS d
            CROSS APPLY
            (
                SELECT
                    CASE
                        WHEN d.Debe > 0 THEN d.Debe
                        ELSE d.Haber
                    END AS ImporteLinea,
                    CASE WHEN @TipoCambio > 0 THEN @TipoCambio ELSE 1 END AS TipoCambioAplicado
            ) AS calc
            ORDER BY
                d.Item ASC;
        END;

        INSERT INTO dbo.COM_CompraDetalle
        (
            IdCompra,
            Item,
            IdPlanCuenta,
            IdTipoAfectacionIGV,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto,
            UsuarioRegistro
        )
        SELECT
            @IdCompraTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.IdTipoAfectacionIGV,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto,
            @UsuarioRegistro
        FROM @DetalleCompra AS d
        ORDER BY
            d.Item ASC;

        COMMIT;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        SELECT
            c.IdCompra,
            c.IdAsiento,
            cd.IdAsiento AS IdAsientoDetraccion,
            c.ImporteTotal,
            c.ImporteDetraccion,
            c.Estado
        FROM dbo.COM_Compra AS c
        LEFT JOIN dbo.COM_CompraDetraccion AS cd
            ON cd.IdCompra = c.IdCompra
        WHERE c.IdCompra = @IdCompraTrabajo;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

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
