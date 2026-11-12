-- Firma: FRANCO LARA | Create date: 10/11/2026 | Consulta negociaciones, facturas elegibles y reporte del lote.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_FactoringConsultar
    @UsuarioId NVARCHAR(128), @Id BIGINT=NULL, @Desde DATE=NULL, @Hasta DATE=NULL,
    @Busqueda NVARCHAR(150)=NULL, @Pagina INT=1, @Elegibles BIT=0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'FACTORING.CONSULTAR')=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para consultar factoring.',1;
        IF @Pagina NOT BETWEEN 1 AND 100000 OR @Desde>@Hasta
            THROW 51601,N'[HYT_VALIDACION] Revisa el rango de fechas y la página.',1;

        SELECT (
            SELECT JSON_QUERY(CASE WHEN @Id IS NOT NULL THEN (
                SELECT lote.LoteId,lote.Financiera,lote.FechaNegociacion,lote.FechaDesembolso,lote.Porcentaje,lote.Moneda,lote.Referencia,lote.Observaciones,
                    JSON_QUERY(COALESCE((
                        SELECT c.ComprobanteId,c.Serie,c.Numero,c.Fecha,c.FechaVencimiento,c.NombreTercero,c.DocumentoTercero,c.Importe
                        FROM dbo.FactoringFacturaHYT AS factura
                        INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=factura.ComprobanteId
                        WHERE factura.LoteId=lote.LoteId ORDER BY c.Fecha,c.ComprobanteId FOR JSON PATH),N'[]')) AS Facturas
                FROM dbo.FactoringLoteHYT AS lote WHERE lote.LoteId=@Id FOR JSON PATH,WITHOUT_ARRAY_WRAPPER) END) AS Lote,
            JSON_QUERY(COALESCE(CASE WHEN @Id IS NULL AND @Elegibles=0 THEN (
                SELECT lote.LoteId,lote.Financiera,lote.FechaNegociacion,lote.FechaDesembolso,lote.Porcentaje,lote.Moneda,
                    resumen.Cantidad,resumen.Importe,COUNT_BIG(1) OVER() AS TotalFilas
                FROM dbo.FactoringLoteHYT AS lote
                CROSS APPLY (SELECT COUNT(1) AS Cantidad,SUM(c.Importe) AS Importe
                    FROM dbo.FactoringFacturaHYT AS factura INNER JOIN dbo.ComprobanteHYT AS c ON c.ComprobanteId=factura.ComprobanteId
                    WHERE factura.LoteId=lote.LoteId) AS resumen
                WHERE lote.FechaNegociacion BETWEEN @Desde AND @Hasta
                    AND (@Busqueda IS NULL OR CHARINDEX(@Busqueda,lote.Financiera)>0 OR CHARINDEX(@Busqueda,lote.Referencia)>0)
                ORDER BY lote.FechaNegociacion DESC,lote.LoteId DESC
                OFFSET (CONVERT(BIGINT,@Pagina)-1)*20 ROWS FETCH NEXT 20 ROWS ONLY FOR JSON PATH) END,N'[]')) AS Lotes,
            JSON_QUERY(COALESCE(CASE WHEN @Elegibles=1 THEN (
                SELECT c.ComprobanteId,c.Serie,c.Numero,c.Fecha,c.FechaVencimiento,c.NombreTercero,c.Moneda,c.Importe
                FROM dbo.ComprobanteHYT AS c
                WHERE c.EsVenta=1 AND c.Tipo='01' AND c.Serie IS NOT NULL AND c.Numero IS NOT NULL AND c.Importe>0
                    AND c.Fecha BETWEEN @Desde AND @Hasta
                    AND (@Busqueda IS NULL OR CHARINDEX(@Busqueda,c.NombreTercero)>0 OR CHARINDEX(@Busqueda,CONCAT(c.Serie,N'-',c.Numero))>0)
                    AND NOT EXISTS(SELECT 1 FROM dbo.FactoringFacturaHYT AS factura WHERE factura.ComprobanteId=c.ComprobanteId)
                ORDER BY c.Fecha DESC,c.ComprobanteId DESC FOR JSON PATH) END,N'[]')) AS Facturas,
            JSON_QUERY(COALESCE(CASE WHEN @Elegibles=1 THEN (
                SELECT banco.Nombre FROM dbo.BancoPeruHYT AS banco ORDER BY banco.Nombre FOR JSON PATH) END,N'[]')) AS Financieras
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_FactoringConsultar TO hyt_app_executor;
GO
