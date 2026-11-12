-- Firma: FRANCO LARA | Create date: 10/11/2026 | Operaciones de provisiones de compras y comprobantes de venta.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_Pago
    @UsuarioId NVARCHAR(128), @Id BIGINT, @EsDetraccion BIT, @Fecha DATE,
    @Importe DECIMAL(18,2), @Referencia NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY

        IF dbo.fn_HYTPermiso(@UsuarioId,N'PAGOS_PROVEEDOR.EDITAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso de edición de pagos.',1;
        IF @@TRANCOUNT<>0 OR @EsDetraccion IS NULL OR @Fecha IS NULL OR @Importe IS NULL OR @Importe<=0 OR NULLIF(LTRIM(RTRIM(@Referencia)),N'') IS NULL
            THROW 51601,N'[HYT_VALIDACION] Completa la fecha, el importe y la referencia del pago.',1;
        BEGIN TRANSACTION;
        DECLARE @Limite DECIMAL(18,2),@Pagado DECIMAL(18,2),@PagoId BIGINT;
        SELECT @Limite=CASE WHEN @EsDetraccion=1 THEN ImporteDetraccion ELSE Importe-DetraccionDocumento END
        FROM dbo.ComprobanteHYT WITH(UPDLOCK,HOLDLOCK) WHERE EsVenta=0 AND Tipo<>'07' AND ComprobanteId=@Id;
        SELECT @PagoId=PagoId FROM dbo.ComprobanteHYTPago WHERE ComprobanteId=@Id AND EsDetraccion=@EsDetraccion AND Referencia=@Referencia AND Fecha=@Fecha AND Importe=@Importe;
        IF @PagoId IS NOT NULL
        BEGIN
            COMMIT TRANSACTION;
            SELECT @PagoId;
            RETURN;
        END;
        SELECT @Pagado=COALESCE(SUM(Importe),0) FROM dbo.ComprobanteHYTPago WHERE ComprobanteId=@Id AND EsDetraccion=@EsDetraccion;
        IF @Limite IS NULL OR @Pagado+@Importe>@Limite
            THROW 51601,N'[HYT_VALIDACION] El importe supera el saldo disponible.',1;
        IF EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPago WHERE ComprobanteId=@Id AND EsDetraccion=@EsDetraccion AND Referencia=@Referencia)
            THROW 51601,N'[HYT_VALIDACION] La referencia ya se usó para otro pago.',1;
        INSERT INTO dbo.ComprobanteHYTPago(ComprobanteId,EsDetraccion,Fecha,Importe,Referencia,UsuarioRegistro)
        VALUES(@Id,@EsDetraccion,@Fecha,@Importe,@Referencia,@UsuarioId);
        SET @PagoId=SCOPE_IDENTITY();
        DECLARE @Registro NVARCHAR(50)=CONVERT(NVARCHAR(50),@PagoId),@Operacion UNIQUEIDENTIFIER=NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'ComprobanteHYTPago',@Registro,N'REGISTRAR';
        COMMIT TRANSACTION;
        SELECT @PagoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_Pago TO hyt_app_executor;
GO
