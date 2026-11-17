-- Firma: FRANCO LARA | Create date: 17/11/2026 | Elimina comprobantes sin pagos ni factoring y sus vínculos, con auditoría y sincronización de OC.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_Eliminar
    @UsuarioId NVARCHAR(128),@EsVenta BIT,@Id BIGINT,@VersionEdicion VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @EsVenta=1 THEN N'FACTURADO.CONSULTAR' ELSE N'PAGOS_PROVEEDOR.CONSULTAR' END)=0
            OR dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @EsVenta=1 THEN N'FACTURADO.EDITAR' ELSE N'PAGOS_PROVEEDOR.EDITAR' END)=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para eliminar el comprobante.',1;
        DECLARE @Version BINARY(8)=TRY_CONVERT(BINARY(8),@VersionEdicion,1);
        IF @Id IS NULL OR @Id<=0 OR LEN(@VersionEdicion)<>18 OR @Version IS NULL
            THROW 51601,N'[HYT_VALIDACION] Vuelve a abrir el registro para eliminarlo.',1;
        BEGIN TRANSACTION;
        DECLARE @Actual BINARY(8);
        SELECT @Actual=CONVERT(BINARY(8),documento.VersionFila)
        FROM dbo.ComprobanteHYT AS documento WITH(UPDLOCK,HOLDLOCK)
        WHERE documento.ComprobanteId=@Id AND documento.EsVenta=@EsVenta;
        IF @Actual IS NULL
        BEGIN
            COMMIT TRANSACTION;
            SELECT @Id;
            RETURN;
        END;
        IF @Actual<>@Version
            THROW 51601,N'[HYT_CONCURRENCIA] El comprobante cambió. Vuelve a abrirlo antes de eliminar.',1;
        IF EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPago AS pago WITH(UPDLOCK,HOLDLOCK) WHERE pago.ComprobanteId=@Id)
            THROW 51601,N'[HYT_VALIDACION] El registro tiene pagos o cancelaciones de detracción. No se puede eliminar.',1;
        IF EXISTS(SELECT 1 FROM dbo.FactoringFacturaHYT AS factura WITH(UPDLOCK,HOLDLOCK) WHERE factura.ComprobanteId=@Id)
            THROW 51601,N'[HYT_VALIDACION] El comprobante pertenece a un lote de factoring. No se puede eliminar.',1;

        DECLARE @Anterior NVARCHAR(MAX)=(
            SELECT documento.ComprobanteId,documento.EsVenta,documento.Tipo,documento.Serie,documento.Numero,
                documento.Fecha,documento.FechaVencimiento,documento.Moneda,documento.ClienteId,documento.ProveedorId,
                documento.OrdenCompraId,documento.DocumentoTercero,documento.NombreTercero,documento.Importe,
                documento.CodigoDetraccion,documento.ImporteDetraccion,documento.DetraccionDocumento,documento.MonedaDetraccion,
                documento.ReferenciaNota,documento.MotivoNota,JSON_QUERY(documento.Detalle) AS Detalle,
                JSON_QUERY((SELECT vinculo.PresupuestoId FROM dbo.ComprobanteHYTPresupuesto AS vinculo
                    WHERE vinculo.ComprobanteId=@Id FOR JSON PATH)) AS Presupuestos,
                JSON_QUERY((SELECT vinculo.CostoProveedorId FROM dbo.ComprobanteHYTCosto AS vinculo
                    WHERE vinculo.ComprobanteId=@Id FOR JSON PATH)) AS Trabajos
            FROM dbo.ComprobanteHYT AS documento WHERE documento.ComprobanteId=@Id
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES
        );
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Operacion UNIQUEIDENTIFIER=NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'ComprobanteHYT',@Registro,N'ELIMINAR',@CambiosJson=@Anterior;
        DELETE FROM dbo.ComprobanteHYTCosto WHERE ComprobanteId=@Id;
        DELETE FROM dbo.ComprobanteHYTPresupuesto WHERE ComprobanteId=@Id;
        DELETE FROM dbo.ComprobanteHYT WHERE ComprobanteId=@Id AND EsVenta=@EsVenta;
        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_Eliminar TO hyt_app_executor;
GO
