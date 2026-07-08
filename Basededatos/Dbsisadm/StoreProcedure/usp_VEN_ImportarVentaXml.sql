-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Importa ventas desde XML SUNAT en estado EN REVISION sin generar asiento contable hasta la provision final.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ImportarVentaXml
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

        DECLARE @IdVentaTrabajo INT;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la venta importada.', 16, 1);
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
            RAISERROR(N'El cliente indicado no existe para la empresa activa.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
              AND c.IdEmpresa = @IdEmpresa
              AND c.ModuloOperacion = 'VEN'
              AND c.Activo = 1
        )
        BEGIN
            RAISERROR(N'La configuracion contable de ventas no existe o esta inactiva.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS t
            WHERE t.CodigoTipoComprobante = @TipoComprobante
              AND t.UsoVentas = 1
              AND t.Estado = 1
        )
        BEGIN
            RAISERROR(N'El tipo de comprobante no esta habilitado para ventas.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.VEN_Venta AS v
            WHERE v.IdEmpresa = @IdEmpresa
              AND v.IdCliente = @IdCliente
              AND v.TipoComprobante = @TipoComprobante
              AND v.Serie = @Serie
              AND v.Numero = @Numero
        )
        BEGIN
            RAISERROR(N'Ya existe una venta con el mismo cliente, tipo, serie y numero.', 16, 1);
        END;

        DECLARE @DetalleVenta TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NULL,
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
            TRY_CONVERT(SMALLINT, d.value('@Item', 'VARCHAR(10)')),
            TRY_CONVERT(INT, NULLIF(d.value('@IdPlanCuenta', 'VARCHAR(20)'), '')),
            TRY_CONVERT(INT, d.value('@IdTipoAfectacionIGV', 'VARCHAR(20)')),
            LEFT(d.value('@Descripcion', 'NVARCHAR(250)'), 250),
            TRY_CONVERT(DECIMAL(18,4), d.value('@Cantidad', 'VARCHAR(30)')),
            TRY_CONVERT(DECIMAL(18,6), d.value('@ValorUnitario', 'VARCHAR(30)')),
            TRY_CONVERT(DECIMAL(18,2), d.value('@ImporteBruto', 'VARCHAR(30)'))
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS x(d);

        IF NOT EXISTS (SELECT 1 FROM @DetalleVenta)
        BEGIN
            RAISERROR(N'La venta importada no contiene lineas de detalle.', 16, 1);
        END;

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
            NULL,
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
            N'EN REVISION',
            @UsuarioRegistro
        );

        SET @IdVentaTrabajo = SCOPE_IDENTITY();

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
        ORDER BY d.Item;

        SELECT
            v.IdVenta,
            v.Estado,
            v.ImporteTotal
        FROM dbo.VEN_Venta AS v
        WHERE v.IdVenta = @IdVentaTrabajo;

    END TRY

    BEGIN CATCH

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
