-- Firma: FRANCO LARA | Create date: 17/11/2026 | Mantiene la referencia vigente de venta por OC y archiva las automáticas eliminadas o anuladas.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Facturado_SincronizarEmision
    @Ordenes NVARCHAR(MAX)=NULL, @Motivo NVARCHAR(1000)=N'Comprobante sin vigencia o desvinculado.'
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Propia BIT=0;
        IF @@TRANCOUNT=0
        BEGIN
            BEGIN TRANSACTION;
            SET @Propia=1;
        END;
        DECLARE @Objetivo TABLE (OrdenCompraId BIGINT PRIMARY KEY,PresupuestoId BIGINT,Estado NVARCHAR(30));
        INSERT INTO @Objetivo(OrdenCompraId,PresupuestoId,Estado)
        SELECT orden.OrdenCompraId,orden.PresupuestoId,orden.Estado
        FROM dbo.OrdenCompraCliente AS orden WITH(UPDLOCK,HOLDLOCK)
        WHERE @Ordenes IS NULL OR orden.OrdenCompraId IN (
            SELECT TRY_CONVERT(BIGINT,item.value) FROM OPENJSON(@Ordenes) AS item
        );

        DECLARE @Facturas TABLE
        (
            OrdenCompraId BIGINT PRIMARY KEY,ComprobanteId BIGINT,Tipo CHAR(2),Serie NVARCHAR(30),Numero NVARCHAR(50),
            Fecha DATE,FechaRegistro DATETIME2(3),UsuarioRegistro NVARCHAR(128)
        );
        INSERT INTO @Facturas(OrdenCompraId,ComprobanteId,Tipo,Serie,Numero,Fecha,FechaRegistro,UsuarioRegistro)
        SELECT objetivo.OrdenCompraId,factura.ComprobanteId,factura.Tipo,factura.Serie,factura.Numero,
            factura.Fecha,factura.FechaRegistro,factura.UsuarioRegistro
        FROM @Objetivo AS objetivo
        OUTER APPLY (
            SELECT TOP (1) vigente.ComprobanteId,vigente.Tipo,vigente.Serie,vigente.Numero,
                vigente.Fecha,vigente.FechaRegistro,vigente.UsuarioRegistro
            FROM (
                SELECT documento.ComprobanteId FROM dbo.ComprobanteHYT AS documento
                WHERE documento.OrdenCompraId=objetivo.OrdenCompraId AND documento.EsVenta=1 AND documento.Tipo IN ('01','03')
                UNION
                SELECT vinculo.ComprobanteId FROM dbo.ComprobanteHYTPresupuesto AS vinculo
                INNER JOIN dbo.ComprobanteHYT AS documento ON documento.ComprobanteId=vinculo.ComprobanteId
                WHERE vinculo.PresupuestoId=objetivo.PresupuestoId AND documento.OrdenCompraId IS NULL
                    AND documento.EsVenta=1 AND documento.Tipo IN ('01','03')
            ) AS candidato
            CROSS APPLY dbo.fn_ComprobanteHYT_FacturaVigente(candidato.ComprobanteId) AS vigente
            ORDER BY vigente.ComprobanteId
        ) AS factura
        WHERE objetivo.Estado<>N'CANCELADA' AND factura.ComprobanteId IS NOT NULL;

        INSERT INTO dbo.ReferenciaComprobanteHistorial(ReferenciaId,OrdenCompraId,ComprobanteId,FechaDocumento,
            CodigoTipoDocumento,Serie,NumeroDocumento,FechaRegistro,UsuarioRegistro,FechaRetiro,MotivoRetiro)
        SELECT referencia.ReferenciaId,referencia.OrdenCompraId,referencia.ComprobanteId,referencia.FechaDocumento,
            referencia.CodigoTipoDocumento,referencia.Serie,referencia.NumeroDocumento,referencia.FechaRegistro,
            referencia.UsuarioRegistro,SYSUTCDATETIME(),@Motivo
        FROM dbo.ReferenciaComprobanteExterno AS referencia
        INNER JOIN @Objetivo AS objetivo ON objetivo.OrdenCompraId=referencia.OrdenCompraId
        LEFT JOIN @Facturas AS factura ON factura.OrdenCompraId=referencia.OrdenCompraId
        WHERE referencia.ComprobanteId IS NOT NULL
            AND (factura.ComprobanteId IS NULL OR factura.ComprobanteId<>referencia.ComprobanteId);

        DELETE referencia FROM dbo.ReferenciaComprobanteExterno AS referencia
        INNER JOIN @Objetivo AS objetivo ON objetivo.OrdenCompraId=referencia.OrdenCompraId
        LEFT JOIN @Facturas AS factura ON factura.OrdenCompraId=referencia.OrdenCompraId
        WHERE referencia.ComprobanteId IS NOT NULL
            AND (factura.ComprobanteId IS NULL OR factura.ComprobanteId<>referencia.ComprobanteId);

        UPDATE referencia SET FechaDocumento=factura.Fecha,CodigoTipoDocumento=factura.Tipo,
            Serie=factura.Serie,NumeroDocumento=factura.Numero
        FROM dbo.ReferenciaComprobanteExterno AS referencia
        INNER JOIN @Facturas AS factura ON factura.OrdenCompraId=referencia.OrdenCompraId
            AND factura.ComprobanteId=referencia.ComprobanteId
        WHERE referencia.FechaDocumento<>factura.Fecha OR referencia.CodigoTipoDocumento<>factura.Tipo
            OR referencia.Serie<>factura.Serie OR referencia.NumeroDocumento<>factura.Numero;

        INSERT INTO dbo.ReferenciaComprobanteExterno(OrdenCompraId,ConfirmacionId,ComprobanteId,
            FechaDocumento,CodigoTipoDocumento,Serie,NumeroDocumento,FechaRegistro,UsuarioRegistro)
        SELECT factura.OrdenCompraId,NULL,factura.ComprobanteId,factura.Fecha,factura.Tipo,factura.Serie,
            factura.Numero,factura.FechaRegistro,factura.UsuarioRegistro
        FROM @Facturas AS factura
        WHERE NOT EXISTS (
            SELECT 1 FROM dbo.ReferenciaComprobanteExterno AS referencia WITH(UPDLOCK,HOLDLOCK)
            WHERE referencia.OrdenCompraId=factura.OrdenCompraId AND referencia.Vigente=1
        );
        IF @Propia=1 COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
