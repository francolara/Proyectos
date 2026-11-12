-- Firma: FRANCO LARA | Create date: 10/11/2026 | Operaciones de provisiones de compras y comprobantes de venta.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Asocia o retira la orden de compra de una venta registrada.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_Vincular
    @UsuarioId NVARCHAR(128), @EsVenta BIT, @Id BIGINT, @Ids NVARCHAR(MAX)=NULL,
    @OperacionId UNIQUEIDENTIFIER=NULL, @PresupuestoId BIGINT=NULL, @Descripcion NVARCHAR(300)=NULL, @Importe DECIMAL(18,2)=NULL,
    @ActualizarOrdenCompra BIT=0,@OrdenCompraId BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @EsVenta=1 THEN N'FACTURADO.EDITAR' ELSE N'PAGOS_PROVEEDOR.EDITAR' END)=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF @@TRANCOUNT<>0 THROW 51601,N'[HYT_VALIDACION] Transacción externa inválida.',1;
        BEGIN TRANSACTION;
        DECLARE @Cliente BIGINT,@Proveedor BIGINT,@Moneda NVARCHAR(3),@Documento NVARCHAR(100),@Total DECIMAL(18,2);
        SELECT @Cliente=ClienteId,@Proveedor=ProveedorId,@Moneda=Moneda,@Documento=CONCAT(Serie,N'-',Numero),@Total=Importe
        FROM dbo.ComprobanteHYT WITH(UPDLOCK,HOLDLOCK) WHERE ComprobanteId=@Id AND EsVenta=@EsVenta;
        IF @Moneda IS NULL THROW 51601,N'[HYT_VALIDACION] No se encontró el comprobante.',1;
        IF @ActualizarOrdenCompra=1
        BEGIN
            IF @EsVenta<>1 OR (@OrdenCompraId IS NOT NULL AND NOT EXISTS (
                SELECT 1 FROM dbo.OrdenCompraCliente AS orden WITH(HOLDLOCK)
                INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=orden.PresupuestoId
                WHERE orden.OrdenCompraId=@OrdenCompraId AND p.ClienteId=@Cliente AND orden.Estado<>N'CANCELADA'))
                THROW 51601,N'[HYT_VALIDACION] Selecciona una orden de compra vigente del cliente.',1;
            UPDATE dbo.ComprobanteHYT SET OrdenCompraId=@OrdenCompraId WHERE ComprobanteId=@Id;
        END
        ELSE IF @EsVenta=1
        BEGIN
            IF NOT EXISTS(SELECT 1 FROM OPENJSON(@Ids)) OR EXISTS(
                SELECT 1 FROM OPENJSON(@Ids) AS elegido LEFT JOIN dbo.Presupuesto AS p ON p.PresupuestoId=TRY_CONVERT(BIGINT,elegido.value)
                WHERE p.PresupuestoId IS NULL OR p.ClienteId<>@Cliente)
                THROW 51601,N'[HYT_VALIDACION] Selecciona presupuestos del cliente.',1;
            INSERT INTO dbo.ComprobanteHYTPresupuesto(ComprobanteId,PresupuestoId)
            SELECT DISTINCT @Id,CONVERT(BIGINT,elegido.value) FROM OPENJSON(@Ids) AS elegido
            WHERE NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPresupuesto AS v WHERE v.ComprobanteId=@Id AND v.PresupuestoId=TRY_CONVERT(BIGINT,elegido.value));
        END
        ELSE
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId,N'COSTOS_PROVEEDOR.EDITAR')=0
                THROW 51600,N'[HYT_PERMISO] Necesita permiso de edición de costos.',1;
            IF @PresupuestoId IS NOT NULL
            BEGIN
                DECLARE @CostoPrevio BIGINT;
                SELECT @CostoPrevio=costo.CostoProveedorId FROM dbo.CostoProveedorPresupuesto AS costo WITH(UPDLOCK,HOLDLOCK)
                WHERE costo.OperacionId=@OperacionId;
                IF @CostoPrevio IS NOT NULL
                BEGIN
                    IF NOT EXISTS(SELECT 1 FROM dbo.CostoProveedorPresupuesto AS costo INNER JOIN dbo.ComprobanteHYTCosto AS v ON v.CostoProveedorId=costo.CostoProveedorId
                        WHERE costo.CostoProveedorId=@CostoPrevio AND v.ComprobanteId=@Id AND costo.PresupuestoId=@PresupuestoId
                            AND costo.Descripcion=@Descripcion AND costo.Importe=@Importe AND costo.UsuarioRegistro=@UsuarioId)
                        THROW 51601,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos.',1;
                    COMMIT TRANSACTION;
                    SELECT @Id;
                    RETURN;
                END;
                IF @Importe IS NULL OR @Importe<=0 OR NULLIF(LTRIM(RTRIM(@Descripcion)),N'') IS NULL
                    OR NOT EXISTS(SELECT 1 FROM dbo.Presupuesto WHERE PresupuestoId=@PresupuestoId)
                    OR @Importe+COALESCE((SELECT SUM(costo.Importe) FROM dbo.ComprobanteHYTCosto AS v INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId=v.CostoProveedorId WHERE v.ComprobanteId=@Id),0)>@Total
                    THROW 51601,N'[HYT_VALIDACION] Revisa el presupuesto, la descripción y el saldo disponible para registrar el costo.',1;
                DECLARE @OperacionCosto UNIQUEIDENTIFIER=COALESCE(@OperacionId,NEWID()),@CostoId BIGINT;
                INSERT INTO dbo.CostoProveedorPresupuesto(PresupuestoId,ProveedorId,Descripcion,Importe,MonedaCodigo,DocumentoSustento,OperacionId,UsuarioRegistro)
                VALUES(@PresupuestoId,@Proveedor,@Descripcion,@Importe,@Moneda,@Documento,@OperacionCosto,@UsuarioId);
                SET @CostoId=SCOPE_IDENTITY();
                INSERT INTO dbo.ComprobanteHYTCosto(ComprobanteId,CostoProveedorId) VALUES(@Id,@CostoId);
                DECLARE @RegistroCosto NVARCHAR(50)=CONVERT(NVARCHAR(50),@CostoId);
                EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@OperacionCosto,N'CostoProveedorPresupuesto',@RegistroCosto,N'REGISTRAR';
            END
            ELSE
            BEGIN
                IF NOT EXISTS(SELECT 1 FROM OPENJSON(@Ids)) OR EXISTS(
                    SELECT 1 FROM OPENJSON(@Ids) AS elegido
                    LEFT JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId=TRY_CONVERT(BIGINT,elegido.value)
                    LEFT JOIN dbo.ComprobanteHYTCosto AS v WITH(UPDLOCK,HOLDLOCK) ON v.CostoProveedorId=costo.CostoProveedorId
                    WHERE costo.CostoProveedorId IS NULL OR costo.ProveedorId<>@Proveedor OR costo.MonedaCodigo<>@Moneda OR (v.ComprobanteId IS NOT NULL AND v.ComprobanteId<>@Id))
                    THROW 51601,N'[HYT_VALIDACION] Selecciona costos disponibles del mismo proveedor y moneda.',1;
                INSERT INTO dbo.ComprobanteHYTCosto(ComprobanteId,CostoProveedorId)
                SELECT DISTINCT @Id,CONVERT(BIGINT,elegido.value) FROM OPENJSON(@Ids) AS elegido
                WHERE NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYTCosto AS v WHERE v.CostoProveedorId=TRY_CONVERT(BIGINT,elegido.value));
                IF (SELECT SUM(costo.Importe) FROM dbo.ComprobanteHYTCosto AS v INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId=v.CostoProveedorId WHERE v.ComprobanteId=@Id)>@Total
                    THROW 51601,N'[HYT_VALIDACION] Los costos vinculados superan el importe de la provisión.',1;
            END
        END;
        DECLARE @Registro NVARCHAR(50)=CONVERT(NVARCHAR(50),@Id),@Operacion UNIQUEIDENTIFIER=NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'ComprobanteHYT',@Registro,N'VINCULAR';
        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_Vincular TO hyt_app_executor;
GO
