-- Firma: FRANCO LARA | Create date: 10/11/2026 | Registra un lote de facturas en una sola operación.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_FactoringGuardar
    @UsuarioId NVARCHAR(128), @Datos NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET XACT_ABORT ON;
        IF dbo.fn_HYTPermiso(@UsuarioId,N'FACTORING.EDITAR')=0 OR dbo.fn_HYTPermiso(@UsuarioId,N'FACTORING.CONSULTAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para registrar factoring.',1;
        IF @Datos IS NULL OR ISJSON(@Datos)<>1 OR @@TRANCOUNT<>0
            THROW 51601,N'[HYT_VALIDACION] Datos del lote inválidos.',1;

        DECLARE @Financiera NVARCHAR(150),@Fecha DATE,@Desembolso DATE,@Porcentaje DECIMAL(7,4),@Referencia NVARCHAR(100),
            @Observaciones NVARCHAR(500),@Operacion UNIQUEIDENTIFIER,@Facturas NVARCHAR(MAX),@LoteId BIGINT,@Moneda NVARCHAR(3);
        SELECT @Financiera=NULLIF(TRIM(Financiera),N''),@Fecha=FechaNegociacion,@Desembolso=FechaDesembolso,@Porcentaje=Porcentaje,
            @Referencia=Referencia,@Observaciones=Observaciones,@Operacion=OperacionId,@Facturas=Facturas
        FROM OPENJSON(@Datos) WITH(Financiera NVARCHAR(150),FechaNegociacion DATE,FechaDesembolso DATE,Porcentaje DECIMAL(7,4),
            Referencia NVARCHAR(100),Observaciones NVARCHAR(500),OperacionId UNIQUEIDENTIFIER,Facturas NVARCHAR(MAX) AS JSON);
        IF @Financiera IS NULL OR @Fecha IS NULL OR @Desembolso<@Fecha OR @Porcentaje IS NULL OR @Porcentaje<=0 OR @Porcentaje>100
            OR @Operacion IS NULL OR (SELECT COUNT(1) FROM OPENJSON(@Facturas)) NOT BETWEEN 1 AND 100
            THROW 51601,N'[HYT_VALIDACION] Completa la financiera, fechas, porcentaje y entre 1 y 100 facturas.',1;
        DECLARE @Seleccion TABLE (ComprobanteId BIGINT PRIMARY KEY);
        IF EXISTS(SELECT 1 FROM OPENJSON(@Facturas) AS elegido WHERE TRY_CONVERT(BIGINT,elegido.value) IS NULL)
            THROW 51601,N'[HYT_VALIDACION] La selección contiene una factura inválida.',1;
        INSERT INTO @Seleccion(ComprobanteId) SELECT DISTINCT CONVERT(BIGINT,elegido.value) FROM OPENJSON(@Facturas) AS elegido;

        DECLARE @Huella VARBINARY(32)=HASHBYTES('SHA2_256',@Datos),@Bloqueo INT;
        BEGIN TRANSACTION;
        EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:Factoring:Registrar',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
        IF @Bloqueo<0 THROW 51601,N'[HYT_VALIDACION] No se pudo registrar el lote. Inténtalo nuevamente.',1;
        SELECT @LoteId=LoteId FROM dbo.FactoringLoteHYT WHERE OperacionId=@Operacion;
        IF @LoteId IS NOT NULL
        BEGIN
            IF NOT EXISTS(SELECT 1 FROM dbo.FactoringLoteHYT WHERE LoteId=@LoteId AND Huella=@Huella AND UsuarioRegistro=@UsuarioId)
                THROW 51601,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos.',1;
            COMMIT TRANSACTION;
            SELECT @LoteId;
            RETURN;
        END;
        IF EXISTS(SELECT 1 FROM @Seleccion AS elegido LEFT JOIN dbo.ComprobanteHYT AS c WITH(UPDLOCK,HOLDLOCK) ON c.ComprobanteId=elegido.ComprobanteId
            LEFT JOIN dbo.FactoringFacturaHYT AS factura ON factura.ComprobanteId=elegido.ComprobanteId
            WHERE c.ComprobanteId IS NULL OR c.EsVenta=0 OR c.Tipo<>'01' OR c.Serie IS NULL OR c.Numero IS NULL OR c.Importe<=0
                OR c.Fecha>@Fecha OR factura.ComprobanteId IS NOT NULL)
            THROW 51601,N'[HYT_VALIDACION] Selecciona facturas emitidas hasta la fecha de negociación y que aún no pertenezcan a otro lote.',1;
        SELECT @Moneda=MIN(c.Moneda) FROM @Seleccion AS elegido INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=elegido.ComprobanteId;
        IF EXISTS(SELECT 1 FROM @Seleccion AS elegido INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=elegido.ComprobanteId WHERE c.Moneda<>@Moneda)
            THROW 51601,N'[HYT_VALIDACION] Agrupa en cada lote facturas de una sola moneda.',1;

        INSERT INTO dbo.FactoringLoteHYT(Financiera,FechaNegociacion,FechaDesembolso,Porcentaje,Moneda,Referencia,Observaciones,OperacionId,Huella,UsuarioRegistro)
        VALUES(@Financiera,@Fecha,@Desembolso,@Porcentaje,@Moneda,@Referencia,@Observaciones,@Operacion,@Huella,@UsuarioId);
        SET @LoteId=SCOPE_IDENTITY();
        INSERT INTO dbo.FactoringFacturaHYT(ComprobanteId,LoteId) SELECT elegido.ComprobanteId,@LoteId FROM @Seleccion AS elegido;
        DECLARE @Registro NVARCHAR(50)=CONVERT(NVARCHAR(50),@LoteId);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'FactoringLoteHYT',@Registro,N'REGISTRAR';
        COMMIT TRANSACTION;
        SELECT @LoteId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_FactoringGuardar TO hyt_app_executor;
GO
