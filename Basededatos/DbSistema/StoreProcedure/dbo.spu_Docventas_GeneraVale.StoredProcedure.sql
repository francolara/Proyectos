-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/10/2026
-- Description:   Genera el vale normal de un documento de venta, obtiene correctamente el tipo de vale desde el motivo y no actualiza stock ni costos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Docventas_GeneraVale
    @IdEmpresa CHAR(2),
    @IdSucursal CHAR(8),
    @IdDocumento CHAR(2),
    @IdSerie CHAR(4),
    @IdDocVentas CHAR(8),
    @IdMotivoTraslado CHAR(8),
    @IdPeriodoInventarioPredeterminado CHAR(8),
    @IdUsuario CHAR(8),
    @PorcentajeIgv DECIMAL(14, 5)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @TransaccionesIniciales INT = @@TRANCOUNT;
    DECLARE @IdVale CHAR(8);
    DECLARE @TipoVale CHAR(1);
    DECLARE @TipoValeAnterior CHAR(1);
    DECLARE @IdValeAnterior CHAR(8);
    DECLARE @IndValeAnterior CHAR(1);
    DECLARE @IdConcepto CHAR(2);
    DECLARE @IdPeriodoInv CHAR(8);
    DECLARE @PrefijoCorrelativo CHAR(4);
    DECLARE @SiguienteCorrelativo INT;
    DECLARE @ValorTotal DECIMAL(14, 5);
    DECLARE @IgvTotal DECIMAL(14, 5);
    DECLARE @PrecioTotal DECIMAL(14, 5);

    DECLARE @Detalle TABLE
    (
        Item BIGINT NOT NULL,
        IdProducto CHAR(8) NOT NULL,
        GlsProducto VARCHAR(185) NOT NULL,
        IdUM CHAR(8) NOT NULL,
        Factor DECIMAL(14, 5) NOT NULL,
        Afecto CHAR(1) NOT NULL,
        Cantidad DECIMAL(14, 5) NOT NULL,
        VVUnit DECIMAL(14, 5) NOT NULL,
        IGVUnit DECIMAL(14, 5) NOT NULL,
        PVUnit DECIMAL(14, 5) NOT NULL,
        TotalVVNeto DECIMAL(14, 5) NOT NULL,
        TotalIGVNeto DECIMAL(14, 5) NOT NULL,
        TotalPVNeto DECIMAL(14, 5) NOT NULL,
        IdMoneda CHAR(8) NOT NULL,
        NumLote VARCHAR(45) NOT NULL,
        FecVencProd VARCHAR(30) NOT NULL
    );

    BEGIN TRY
        IF @TransaccionesIniciales = 0
            BEGIN TRANSACTION;
        ELSE
            SAVE TRANSACTION GeneraValeDocVentas;

        SELECT
            @IndValeAnterior = ISNULL(DV.IndVale, 'N'),
            @IdValeAnterior = ISNULL(DV.IdValesCab, ''),
            @TipoValeAnterior = ISNULL(DV.TipoVale, '')
        FROM dbo.DocVentas AS DV WITH (UPDLOCK, HOLDLOCK)
        WHERE DV.IdEmpresa = @IdEmpresa
          AND DV.IdSucursal = @IdSucursal
          AND DV.IdDocumento = @IdDocumento
          AND DV.IdSerie = @IdSerie
          AND DV.IdDocVentas = @IdDocVentas;

        IF @@ROWCOUNT = 0
            RAISERROR ('No se encontro el documento de venta para generar el vale.', 16, 1);

        SELECT
            @TipoVale = CASE
                WHEN @IdDocumento IN ('01', '03', '12', '90', '91', '08', '86')
                    THEN ISNULL(NULLIF(LTRIM(RTRIM(MT.IndTipoVale)), ''), 'S')
                WHEN @IdDocumento IN ('07', '89') THEN 'I'
                ELSE 'S'
            END
        FROM dbo.MotivosTraslados AS MT
        WHERE MT.IdMotivoTraslado = @IdMotivoTraslado
          AND MT.IdDocumento = @IdDocumento;

        IF @TipoVale IS NULL OR LTRIM(RTRIM(@TipoVale)) = ''
            SET @TipoVale = CASE WHEN @IdDocumento IN ('07', '89') THEN 'I' ELSE 'S' END;

        IF @IdDocumento IN ('07', '89')
        BEGIN
            SELECT @IdConcepto = ISNULL(MN.IdConcepto, '')
            FROM dbo.MotivosNCD AS MN
            WHERE MN.IdMotivoNCD = @IdMotivoTraslado;
        END
        ELSE
        BEGIN
            SELECT @IdConcepto = ISNULL(MT.IdConcepto, '')
            FROM dbo.MotivosTraslados AS MT
            WHERE MT.IdMotivoTraslado = @IdMotivoTraslado
              AND MT.IdDocumento = @IdDocumento;
        END;

        IF @IdConcepto IS NULL OR LTRIM(RTRIM(@IdConcepto)) = ''
            SET @IdConcepto = '22';

        SELECT TOP (1) @IdPeriodoInv = PI.IdPeriodoInv
        FROM dbo.PeriodosInv AS PI
        INNER JOIN dbo.DocVentas AS DV
            ON DV.IdEmpresa = @IdEmpresa
           AND DV.IdSucursal = @IdSucursal
           AND DV.IdDocumento = @IdDocumento
           AND DV.IdSerie = @IdSerie
           AND DV.IdDocVentas = @IdDocVentas
        WHERE PI.IdEmpresa = @IdEmpresa
          AND PI.IdSucursal = @IdSucursal
          AND YEAR(PI.FecInicio) = YEAR(DV.FecEmision)
        ORDER BY PI.FecInicio;

        SET @IdPeriodoInv = ISNULL(NULLIF(@IdPeriodoInv, ''), @IdPeriodoInventarioPredeterminado);

        IF @IndValeAnterior = 'S' AND LTRIM(RTRIM(@IdValeAnterior)) <> ''
            SET @IdVale = @IdValeAnterior;
        ELSE
        BEGIN
            SET @PrefijoCorrelativo = RIGHT(CONVERT(CHAR(4), YEAR(GETDATE())), 2) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(GETDATE())), 2);

            SELECT @SiguienteCorrelativo = ISNULL(MAX(TRY_CONVERT(INT, RIGHT(VC.IdValesCab, 4))), 0) + 1
            FROM dbo.ValesCab AS VC WITH (UPDLOCK, HOLDLOCK)
            WHERE VC.IdEmpresa = @IdEmpresa
              AND VC.TipoVale = @TipoVale
              AND LEFT(VC.IdValesCab, 4) = @PrefijoCorrelativo;

            SET @IdVale = @PrefijoCorrelativo + RIGHT('0000' + CONVERT(VARCHAR(4), @SiguienteCorrelativo), 4);
        END;

        IF @IndValeAnterior = 'S' AND LTRIM(RTRIM(@IdValeAnterior)) <> ''
        BEGIN
            DELETE FROM dbo.ValesDet
            WHERE IdEmpresa = @IdEmpresa
              AND IdSucursal = @IdSucursal
              AND IdValesCab = @IdValeAnterior
              AND TipoVale = @TipoValeAnterior;

            DELETE FROM dbo.DocReferencia
            WHERE IdEmpresa = @IdEmpresa
              AND IdSucursal = @IdSucursal
              AND TipoDocOrigen = CASE WHEN @TipoValeAnterior = 'I' THEN '88' ELSE '99' END
              AND NumDocOrigen = @IdValeAnterior
              AND SerieDocOrigen = '000';

            DELETE FROM dbo.ValesCab
            WHERE IdEmpresa = @IdEmpresa
              AND IdSucursal = @IdSucursal
              AND IdValesCab = @IdValeAnterior
              AND TipoVale = @TipoValeAnterior;
        END;

        INSERT INTO @Detalle
        (
            Item, IdProducto, GlsProducto, IdUM, Factor, Afecto, Cantidad,
            VVUnit, IGVUnit, PVUnit, TotalVVNeto, TotalIGVNeto, TotalPVNeto,
            IdMoneda, NumLote, FecVencProd
        )
        SELECT
            DD.Item,
            DD.IdProducto,
            DD.GlsProducto,
            DD.IdUM,
            DD.Factor,
            DD.Afecto,
            DD.Cantidad,
            CASE WHEN @TipoVale = 'I' THEN DD.VVUnit ELSE 0 END,
            CASE WHEN @TipoVale = 'I' THEN ROUND(DD.VVUnit * @PorcentajeIgv, 5) ELSE 0 END,
            CASE WHEN @TipoVale = 'I' THEN ROUND(DD.VVUnit * (1 + @PorcentajeIgv), 5) ELSE 0 END,
            CASE WHEN @TipoVale = 'I' THEN ROUND(DD.VVUnit * DD.Cantidad, 5) ELSE 0 END,
            CASE WHEN @TipoVale = 'I' THEN ROUND(DD.VVUnit * DD.Cantidad * @PorcentajeIgv, 5) ELSE 0 END,
            CASE WHEN @TipoVale = 'I' THEN ROUND(DD.VVUnit * DD.Cantidad * (1 + @PorcentajeIgv), 5) ELSE 0 END,
            DD.IdMoneda,
            DD.NumLote,
            DD.FecVencProd
        FROM dbo.DocVentasDet AS DD
        INNER JOIN dbo.Productos AS P
            ON P.IdEmpresa = DD.IdEmpresa
           AND P.IdProducto = DD.IdProducto
        WHERE DD.IdEmpresa = @IdEmpresa
          AND DD.IdSucursal = @IdSucursal
          AND DD.IdDocumento = @IdDocumento
          AND DD.IdSerie = @IdSerie
          AND DD.IdDocVentas = @IdDocVentas
          AND P.IdTipoProducto <> '06002';

        IF NOT EXISTS (SELECT 1 FROM @Detalle)
        BEGIN
            UPDATE dbo.DocVentas
            SET TipoVale = '', IdValesCab = '', IndVale = 'N'
            WHERE IdEmpresa = @IdEmpresa
              AND IdSucursal = @IdSucursal
              AND IdDocumento = @IdDocumento
              AND IdSerie = @IdSerie
              AND IdDocVentas = @IdDocVentas;

            IF @TransaccionesIniciales = 0
                COMMIT TRANSACTION;

            SELECT CAST(0 AS BIT) AS IndValeGenerado, '' AS IdValesCab, '' AS TipoVale;
            RETURN;
        END;

        SELECT
            @ValorTotal = SUM(D.TotalVVNeto),
            @IgvTotal = SUM(D.TotalIGVNeto),
            @PrecioTotal = SUM(D.TotalPVNeto)
        FROM @Detalle AS D;

        INSERT INTO dbo.ValesCab
        (
            IdValesCab, TipoVale, FechaEmision, ValorTotal, IgvTotal, PrecioTotal,
            IdProvCliente, IdConcepto, IdAlmacen, IdMoneda, GlsDocReferencia,
            TipoCambio, IdEmpresa, IdSucursal, IdPeriodoInv, IdCentroCosto,
            FechaRegistro, IdUsuarioRegistro
        )
        SELECT
            @IdVale, @TipoVale, DV.FecEmision, @ValorTotal, @IgvTotal, @PrecioTotal,
            DV.IdPerCliente, @IdConcepto, DV.IdAlmacen, DV.IdMoneda,
            D.AbreDocumento + ' ' + @IdSerie + '-' + @IdDocVentas,
            DV.TipoCambio, @IdEmpresa, @IdSucursal, @IdPeriodoInv, DV.IdCentroCosto,
            GETDATE(), @IdUsuario
        FROM dbo.DocVentas AS DV
        INNER JOIN dbo.Documentos AS D
            ON D.IdDocumento = DV.IdDocumento
        WHERE DV.IdEmpresa = @IdEmpresa
          AND DV.IdSucursal = @IdSucursal
          AND DV.IdDocumento = @IdDocumento
          AND DV.IdSerie = @IdSerie
          AND DV.IdDocVentas = @IdDocVentas;

        INSERT INTO dbo.ValesDet
        (
            TipoVale, IdValesCab, Item, IdProducto, GlsProducto, IdUM, Factor, Afecto,
            Cantidad, VVUnit, IGVUnit, PVUnit, TotalVVNeto, TotalIGVNeto, TotalPVNeto,
            IdMoneda, IdEmpresa, IdSucursal, NumLote, FecVencProd
        )
        SELECT
            @TipoVale, @IdVale, D.Item, D.IdProducto, D.GlsProducto, D.IdUM, D.Factor, D.Afecto,
            D.Cantidad, D.VVUnit, D.IGVUnit, D.PVUnit, D.TotalVVNeto, D.TotalIGVNeto, D.TotalPVNeto,
            D.IdMoneda, @IdEmpresa, @IdSucursal, D.NumLote, D.FecVencProd
        FROM @Detalle AS D;

        INSERT INTO dbo.DocReferencia
        (
            TipoDocOrigen, NumDocOrigen, SerieDocOrigen,
            TipoDocReferencia, NumDocReferencia, SerieDocReferencia,
            Item, IdEmpresa, IdSucursal
        )
        VALUES
        (
            CASE WHEN @TipoVale = 'I' THEN '88' ELSE '99' END,
            @IdVale, '000', @IdDocumento, @IdDocVentas, @IdSerie,
            1, @IdEmpresa, @IdSucursal
        );

        UPDATE dbo.DocVentas
        SET TipoVale = @TipoVale,
            IdValesCab = @IdVale,
            IndVale = 'S'
        WHERE IdEmpresa = @IdEmpresa
          AND IdSucursal = @IdSucursal
          AND IdDocumento = @IdDocumento
          AND IdSerie = @IdSerie
          AND IdDocVentas = @IdDocVentas;

        IF @TransaccionesIniciales = 0
            COMMIT TRANSACTION;

        SELECT CAST(1 AS BIT) AS IndValeGenerado, @IdVale AS IdValesCab, @TipoVale AS TipoVale;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
        BEGIN
            IF @TransaccionesIniciales = 0
                ROLLBACK TRANSACTION;
            ELSE IF XACT_STATE() = 1
                ROLLBACK TRANSACTION GeneraValeDocVentas;
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
