-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Importa compras desde XML SUNAT en estado EN REVISION sin generar asiento contable ni documentos pendientes derivados y detalla el documento existente cuando detecta duplicados.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Mejora la validacion de duplicados en la importacion XML de compras informando el IdCompra, fecha y estado del comprobante ya existente.

CREATE OR ALTER PROCEDURE dbo.usp_COM_ImportarCompraXml
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
    @ExoneracionRenta4ta BIT = 0,
    @PorcentajeRetencion DECIMAL(7,4) = 0,
    @Retencion DECIMAL(18,2) = 0,
    @TieneDetraccion BIT = 0,
    @IdDetraccionSunat INT = NULL,
    @PorcentajeDetraccion DECIMAL(7,4) = 0,
    @ImporteDetraccion DECIMAL(18,2) = 0,
    @TienePercepcion BIT = 0,
    @IdTipoPercepcion INT = NULL,
    @PorcentajePercepcion DECIMAL(7,4) = 0,
    @BasePercepcion DECIMAL(18,2) = 0,
    @ImportePercepcion DECIMAL(18,2) = 0,
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCompraTrabajo INT;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la compra importada.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Proveedor AS p
            WHERE p.IdProveedor = @IdProveedor
              AND p.IdEmpresa = @IdEmpresa
              AND p.Estado = 1
        )
        BEGIN
            RAISERROR(N'El proveedor indicado no existe para la empresa activa.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
              AND c.IdEmpresa = @IdEmpresa
              AND c.ModuloOperacion = 'COM'
              AND c.Activo = 1
        )
        BEGIN
            RAISERROR(N'La configuracion contable de compras no existe o esta inactiva.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_TipoComprobante AS t
            WHERE t.CodigoTipoComprobante = @TipoComprobante
              AND t.UsoCompras = 1
              AND t.Estado = 1
        )
        BEGIN
            RAISERROR(N'El tipo de comprobante no esta habilitado para compras.', 16, 1);
        END;

        DECLARE @IdCompraExistente INT;
        DECLARE @FechaEmisionExistente DATE;
        DECLARE @FechaEmisionExistenteTexto VARCHAR(10);
        DECLARE @EstadoExistente NVARCHAR(20);
        DECLARE @EstadoExistenteTexto NVARCHAR(20);

        SELECT TOP (1)
            @IdCompraExistente = c.IdCompra,
            @FechaEmisionExistente = c.FechaEmision,
            @EstadoExistente = c.Estado
        FROM dbo.COM_Compra AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.IdProveedor = @IdProveedor
          AND c.TipoComprobante = @TipoComprobante
          AND c.Serie = @Serie
          AND c.Numero = @Numero
        ORDER BY c.IdCompra DESC;

        IF @IdCompraExistente IS NOT NULL
        BEGIN
            SET @FechaEmisionExistenteTexto = CONVERT(VARCHAR(10), @FechaEmisionExistente, 103);
            SET @EstadoExistenteTexto = ISNULL(@EstadoExistente, N'SIN ESTADO');

            RAISERROR(
                N'Ya existe una compra con el mismo proveedor, tipo, serie y numero. IdCompra=%d, Fecha=%s, Estado=%s.',
                16,
                1,
                @IdCompraExistente,
                @FechaEmisionExistenteTexto,
                @EstadoExistenteTexto);
        END;

        DECLARE @DetalleCompra TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NULL,
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
            TRY_CONVERT(SMALLINT, d.value('@Item', 'VARCHAR(10)')),
            TRY_CONVERT(INT, NULLIF(d.value('@IdPlanCuenta', 'VARCHAR(20)'), '')),
            TRY_CONVERT(INT, d.value('@IdTipoAfectacionIGV', 'VARCHAR(20)')),
            LEFT(d.value('@Descripcion', 'NVARCHAR(250)'), 250),
            TRY_CONVERT(DECIMAL(18,4), d.value('@Cantidad', 'VARCHAR(30)')),
            TRY_CONVERT(DECIMAL(18,6), d.value('@ValorUnitario', 'VARCHAR(30)')),
            TRY_CONVERT(DECIMAL(18,2), d.value('@ImporteBruto', 'VARCHAR(30)'))
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS x(d);

        IF NOT EXISTS (SELECT 1 FROM @DetalleCompra)
        BEGIN
            RAISERROR(N'La compra importada no contiene lineas de detalle.', 16, 1);
        END;

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
            ExoneracionRenta4ta,
            PorcentajeRetencion,
            Retencion,
            TieneDetraccion,
            IdDetraccionSunat,
            PorcentajeDetraccion,
            ImporteDetraccion,
            TienePercepcion,
            IdTipoPercepcion,
            PorcentajePercepcion,
            BasePercepcion,
            ImportePercepcion,
            Observacion,
            Estado,
            UsuarioRegistro
        )
        VALUES
        (
            @IdEmpresa,
            @IdProveedor,
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
            @ImporteTotal - @ImporteDetraccion,
            @ExoneracionRenta4ta,
            @PorcentajeRetencion,
            @Retencion,
            @TieneDetraccion,
            @IdDetraccionSunat,
            @PorcentajeDetraccion,
            @ImporteDetraccion,
            @TienePercepcion,
            @IdTipoPercepcion,
            @PorcentajePercepcion,
            @BasePercepcion,
            @ImportePercepcion,
            @Observacion,
            N'EN REVISION',
            @UsuarioRegistro
        );

        SET @IdCompraTrabajo = SCOPE_IDENTITY();

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
        ORDER BY d.Item;

        SELECT
            c.IdCompra,
            c.Estado,
            c.ImporteTotal
        FROM dbo.COM_Compra AS c
        WHERE c.IdCompra = @IdCompraTrabajo;

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
