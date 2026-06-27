-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza una venta y genera su asiento automatico segun la configuracion contable.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/06/2026
-- Description:   Alinea ventas con compras agregando subtotal, total exonerado, total inafecto, ICBPER interno y detalle con cuenta contable y afectacion IGV.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Usa cuenta del detalle, cuenta por documento/impuesto, valida cuentas activas con movimiento, separa RUC/DNI, tipo comprobante, serie y numero en el asiento e informa fecha de emision en la cabecera contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Inicializa y mantiene el saldo del comprobante de venta igual al importe total al registrar o editar la provision.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Unifica el estado de la provision de venta a PROVISIONADO.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Guarda tipo documento por codigo en ventas y calcula equivalencias en soles y dolares por linea del asiento generado.

CREATE OR ALTER PROCEDURE dbo.usp_VEN_GuardarVentaConAsiento
    @IdVenta INT = NULL,
    @IdEmpresa INT,
    @IdCliente INT,
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
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdVentaTrabajo INT
        DECLARE @IdAsientoTrabajo INT
        DECLARE @IdOrigen INT
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), YEAR(@FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(@FechaContabilizacion)), 2)
        DECLARE @Ejercicio SMALLINT = YEAR(@FechaContabilizacion)
        DECLARE @Mes TINYINT = MONTH(@FechaContabilizacion)
        DECLARE @NumeroAsiento INT
        DECLARE @GlosaAsiento NVARCHAR(500)
        DECLARE @TotalDebe DECIMAL(18,2)
        DECLARE @TotalHaber DECIMAL(18,2)
        DECLARE @TotalExoneradoDetalle DECIMAL(18,2)
        DECLARE @TotalInafectoDetalle DECIMAL(18,2)
        DECLARE @EstadoConfiguracion BIT
        DECLARE @GeneraAsientoAutomatico BIT
        DECLARE @IdTipoComprobanteTrabajo INT
        DECLARE @CodigoMoneda VARCHAR(10)
        DECLARE @IdCuentaDocumento INT
        DECLARE @IdCuentaIgv INT
        DECLARE @IdCuentaIsc INT
        DECLARE @IdCuentaIcbper INT
        DECLARE @IdCuentaOtros INT
        DECLARE @NumeroDocumentoCliente VARCHAR(20)
        DECLARE @DescripcionTipoComprobante NVARCHAR(150)

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
            RAISERROR(N'Los montos de la venta no pueden ser negativos.', 16, 1);
        END;

        IF @ImporteTotal <> (@BaseImponible + @Igv)
        BEGIN
            RAISERROR(N'El importe total debe ser igual a la suma del subtotal e IGV.', 16, 1);
        END;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la venta.', 16, 1);
        END;

        SELECT
            @NumeroDocumentoCliente = per.NumeroDocumento
        FROM dbo.ADM_Cliente AS c
        INNER JOIN dbo.ADM_Persona AS per
            ON per.IdPersona = c.IdPersona
        WHERE c.IdCliente = @IdCliente
          AND c.IdEmpresa = @IdEmpresa
          AND c.Estado = 1
          AND per.IdEmpresa = @IdEmpresa
          AND per.Estado = 1;

        IF @NumeroDocumentoCliente IS NULL
        BEGIN
            RAISERROR(N'El cliente seleccionado no existe o no pertenece a la empresa.', 16, 1);
        END;

        SELECT
            @IdTipoComprobanteTrabajo = t.IdTipoComprobante,
            @DescripcionTipoComprobante = t.Descripcion
        FROM dbo.ADM_TipoComprobante AS t
        WHERE t.CodigoTipoComprobante = @TipoComprobante
          AND t.UsoVentas = 1
          AND t.Estado = 1;

        IF @IdTipoComprobanteTrabajo IS NULL
        BEGIN
            RAISERROR(N'El tipo de comprobante no existe o no esta habilitado para ventas.', 16, 1);
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

        SELECT
            @IdOrigen = c.IdOrigen,
            @EstadoConfiguracion = c.Activo,
            @GeneraAsientoAutomatico = c.GeneraAsientoAutomatico
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
          AND c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'VEN';

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'La configuracion contable indicada no existe para ventas en la empresa activa.', 16, 1);
        END;

        IF @EstadoConfiguracion = 0
        BEGIN
            RAISERROR(N'La configuracion contable seleccionada esta inactiva.', 16, 1);
        END;

        IF @GeneraAsientoAutomatico = 0
        BEGIN
            RAISERROR(N'La configuracion seleccionada no esta habilitada para generar asiento automatico.', 16, 1);
        END;

        DECLARE @DetalleVenta TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            IdTipoAfectacionIGV INT NOT NULL,
            Descripcion NVARCHAR(250) NOT NULL,
            Cantidad DECIMAL(18,4) NOT NULL,
            ValorUnitario DECIMAL(18,6) NOT NULL,
            ImporteBruto DECIMAL(18,2) NOT NULL
        );

        INSERT INTO @DetalleVenta
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
            FROM @DetalleVenta
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos una linea en la venta.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @DetalleVenta AS d
            WHERE d.Item < 1
               OR d.IdPlanCuenta <= 0
               OR d.IdTipoAfectacionIGV <= 0
               OR d.Cantidad <= 0
               OR d.ValorUnitario < 0
               OR d.ImporteBruto < 0
        )
        BEGIN
            RAISERROR(N'El detalle de la venta contiene valores no validos.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.Item
            FROM @DetalleVenta AS d
            GROUP BY
                d.Item
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten items duplicados en el detalle de la venta.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.IdPlanCuenta
            FROM @DetalleVenta AS d
            LEFT JOIN dbo.CON_PlanCuenta AS pc
                ON pc.IdPlanCuenta = d.IdPlanCuenta
               AND pc.IdEmpresa = @IdEmpresa
               AND pc.Estado = 1
               AND pc.AceptaMovimiento = 1
            WHERE pc.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR(N'La cuenta contable seleccionada en el detalle de la venta no es valida. Verifique que pertenezca a la empresa, este activa y acepte movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.IdTipoAfectacionIGV
            FROM @DetalleVenta AS d
            LEFT JOIN dbo.CON_TipoAfectacionIGV AS a
                ON a.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV
               AND a.Estado = 1
            WHERE a.IdTipoAfectacionIGV IS NULL
        )
        BEGIN
            RAISERROR(N'El detalle contiene un tipo de afectacion IGV no valido para la venta.', 16, 1);
        END;

        SELECT
            @TotalExoneradoDetalle = ISNULL(SUM(CASE WHEN a.CodigoSunat LIKE '2%' THEN d.ImporteBruto ELSE 0 END), 0),
            @TotalInafectoDetalle = ISNULL(SUM(CASE WHEN a.CodigoSunat LIKE '3%' THEN d.ImporteBruto ELSE 0 END), 0)
        FROM @DetalleVenta AS d
        INNER JOIN dbo.CON_TipoAfectacionIGV AS a
            ON a.IdTipoAfectacionIGV = d.IdTipoAfectacionIGV;

        IF @TotalExonerado <> @TotalExoneradoDetalle
        BEGIN
            RAISERROR(N'El total exonerado no coincide con el detalle de la venta.', 16, 1);
        END;

        IF @TotalInafecto <> @TotalInafectoDetalle
        BEGIN
            RAISERROR(N'El total inafecto no coincide con el detalle de la venta.', 16, 1);
        END;

        SELECT
            @IdCuentaDocumento = CASE
                WHEN @CodigoMoneda = 'USD' THEN COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdCuentaVentaDolares END, t.IdCuentaVentaDolares)
                ELSE COALESCE(CASE WHEN ISNULL(cfg.Activo, 1) = 1 THEN cfg.IdCuentaVentaSoles END, t.IdCuentaVentaSoles)
            END
        FROM dbo.ADM_TipoComprobante AS t
        LEFT JOIN dbo.CON_DocumentoConfiguracionEmpresa AS cfg
            ON cfg.IdTipoComprobante = t.IdTipoComprobante
           AND cfg.IdEmpresa = @IdEmpresa
        WHERE t.IdTipoComprobante = @IdTipoComprobanteTrabajo;

        IF @IdCuentaDocumento IS NULL
        BEGIN
            RAISERROR(N'No existe una cuenta contable configurada para ventas en el tipo de comprobante y moneda seleccionados.', 16, 1);
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
            RAISERROR(N'La cuenta contable configurada para el documento de venta no pertenece a la empresa, no esta activa o no acepta movimiento.', 16, 1);
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

        DECLARE @AsientoDetalle TABLE
        (
            Item SMALLINT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL
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
            @IdCuentaDocumento,
            @ImporteTotal,
            0,
            CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Documento')
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
            0 AS Debe,
            SUM(d.ImporteBruto) AS Haber,
            CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Detalle')
        FROM @DetalleVenta AS d
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
                0,
                @Igv,
                CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / IGV')
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
                0,
                @Isc,
                CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ISC')
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
                0,
                @Icbper,
                CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ICBPER')
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
                0,
                @OtrosTributos,
                CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / Otros tributos')
            );
        END;

        DELETE FROM @AsientoDetalle
        WHERE Debe = 0
          AND Haber = 0;

        IF NOT EXISTS
        (
            SELECT 1
            FROM @AsientoDetalle
        )
        BEGIN
            RAISERROR(N'No se pudieron generar lineas contables para la venta con la configuracion de documento, impuestos y detalle.', 16, 1);
        END;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @AsientoDetalle AS d;

        IF @TotalDebe <> @TotalHaber
        BEGIN
            RAISERROR(N'La configuracion contable de ventas no genera un asiento cuadrado para los importes ingresados.', 16, 1);
        END;

        SET @GlosaAsiento = CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero);

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

        BEGIN TRAN;

        IF @IdVenta IS NULL
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

            INSERT INTO dbo.VEN_Venta
            (
                IdEmpresa,
                IdCliente,
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
                Observacion,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdCliente,
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
                @ImporteTotal,
                @Observacion,
                N'PROVISIONADO',
                @UsuarioRegistro
            );

            SET @IdVentaTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT
                @IdVentaTrabajo = v.IdVenta,
                @IdAsientoTrabajo = v.IdAsiento
            FROM dbo.VEN_Venta AS v
            WHERE v.IdVenta = @IdVenta
              AND v.IdEmpresa = @IdEmpresa;

            IF @IdVentaTrabajo IS NULL
            BEGIN
                RAISERROR(N'La venta indicada no existe para la empresa activa.', 16, 1);
            END;

            UPDATE dbo.VEN_Venta
            SET IdCliente = @IdCliente,
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
                Saldo = @ImporteTotal,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdVenta = @IdVentaTrabajo;

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

            DELETE FROM dbo.VEN_VentaDetalle
            WHERE IdVenta = @IdVentaTrabajo;
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
            IdCliente,
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
            @NumeroDocumentoCliente,
            @Serie,
            @IdCliente,
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

        INSERT INTO dbo.VEN_VentaDetalle
        (
            IdVenta,
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
            @IdVentaTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.IdTipoAfectacionIGV,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto,
            @UsuarioRegistro
        FROM @DetalleVenta AS d
        ORDER BY
            d.Item ASC;

        COMMIT;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        SELECT
            v.IdVenta,
            v.IdAsiento,
            v.ImporteTotal,
            v.Estado
        FROM dbo.VEN_Venta AS v
        WHERE v.IdVenta = @IdVentaTrabajo;

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
