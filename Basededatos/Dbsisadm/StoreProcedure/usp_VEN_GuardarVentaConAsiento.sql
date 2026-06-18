-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza una venta y genera su asiento automatico segun la configuracion contable.
-- =============================================

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
        DECLARE @EstadoConfiguracion BIT
        DECLARE @GeneraAsientoAutomatico BIT

        IF @BaseImponible < 0
           OR @Igv < 0
           OR @Isc < 0
           OR @OtrosTributos < 0
           OR @Redondeo < 0
           OR @ImporteTotal < 0
        BEGIN
            RAISERROR(N'Los montos de la venta no pueden ser negativos.', 16, 1);
        END;

        IF @ImporteTotal <> (@BaseImponible + @Igv + @Isc + @OtrosTributos + @Redondeo)
        BEGIN
            RAISERROR(N'El importe total debe ser igual a la suma de bruto, IGV, ISC, otros tributos y redondeo.', 16, 1);
        END;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la venta.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Cliente AS c
            WHERE c.IdCliente = @IdCliente
              AND c.IdEmpresa = @IdEmpresa
              AND c.Estado = 1
        )
        BEGIN
            RAISERROR(N'El cliente seleccionado no existe o no pertenece a la empresa.', 16, 1);
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
            Descripcion NVARCHAR(250) NOT NULL,
            Cantidad DECIMAL(18,4) NOT NULL,
            ValorUnitario DECIMAL(18,6) NOT NULL,
            ImporteBruto DECIMAL(18,2) NOT NULL
        );

        INSERT INTO @DetalleVenta
        (
            Item,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto
        )
        SELECT
            T.N.value('@Item', 'smallint'),
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
        SELECT
            d.IdPlanCuenta,
            CASE d.NaturalezaMovimiento
                WHEN 'D' THEN
                    CASE d.ComponenteContable
                        WHEN 'BRUTO' THEN @BaseImponible
                        WHEN 'IGV' THEN @Igv
                        WHEN 'ISC' THEN @Isc
                        WHEN 'OTROS' THEN @OtrosTributos
                        WHEN 'REDONDEO' THEN @Redondeo
                        WHEN 'TOTAL' THEN @ImporteTotal
                        ELSE 0
                    END
                ELSE 0
            END AS Debe,
            CASE d.NaturalezaMovimiento
                WHEN 'H' THEN
                    CASE d.ComponenteContable
                        WHEN 'BRUTO' THEN @BaseImponible
                        WHEN 'IGV' THEN @Igv
                        WHEN 'ISC' THEN @Isc
                        WHEN 'OTROS' THEN @OtrosTributos
                        WHEN 'REDONDEO' THEN @Redondeo
                        WHEN 'TOTAL' THEN @ImporteTotal
                        ELSE 0
                    END
                ELSE 0
            END AS Haber,
            CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ', d.ComponenteContable) AS GlosaDetalle
        FROM dbo.CON_ConfiguracionContabilizacionDetalle AS d
        WHERE d.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
          AND d.Activo = 1;

        DELETE FROM @AsientoDetalle
        WHERE Debe = 0
          AND Haber = 0;

        IF NOT EXISTS
        (
            SELECT 1
            FROM @AsientoDetalle
        )
        BEGIN
            RAISERROR(N'La configuracion seleccionada no genera lineas contables con los importes de la venta.', 16, 1);
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
                @FechaContabilizacion,
                @GlosaAsiento,
                @IdMoneda,
                @TipoCambio,
                @TotalDebe,
                @TotalHaber,
                N'FACTURADO',
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
                Igv,
                Isc,
                OtrosTributos,
                Redondeo,
                ImporteTotal,
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
                @Igv,
                @Isc,
                @OtrosTributos,
                @Redondeo,
                @ImporteTotal,
                @Observacion,
                N'FACTURADO',
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
                Igv = @Igv,
                Isc = @Isc,
                OtrosTributos = @OtrosTributos,
                Redondeo = @Redondeo,
                ImporteTotal = @ImporteTotal,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdVenta = @IdVentaTrabajo;

            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.CON_Asiento
            SET FechaAsiento = @FechaContabilizacion,
                Glosa = @GlosaAsiento,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebe,
                TotalHaber = @TotalHaber,
                Estado = N'FACTURADO',
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
            IdCliente,
            Debe,
            Haber,
            ReferenciaLinea,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.GlosaDetalle,
            @IdCliente,
            d.Debe,
            d.Haber,
            CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
            @UsuarioRegistro
        FROM @AsientoDetalle AS d
        ORDER BY
            d.Item ASC;

        INSERT INTO dbo.VEN_VentaDetalle
        (
            IdVenta,
            Item,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto,
            UsuarioRegistro
        )
        SELECT
            @IdVentaTrabajo,
            d.Item,
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
