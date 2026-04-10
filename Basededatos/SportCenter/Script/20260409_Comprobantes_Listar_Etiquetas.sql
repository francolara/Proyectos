/*
Firma: Codex - 09/04/2026
Descripcion: Actualiza Sp_Comprobantes_Listar con filtro + paginacion backend y total de registros para listado de comprobantes.
*/
/*
Firma: Codex - 10/04/2026
Descripcion: Agrega filtro por tipo de documento (codigo SUNAT) al listado de comprobantes por negocio.
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Comprobantes_Listar]
    @NegocioId INT,
    @SedeId INT = NULL,
    @Buscar NVARCHAR(120) = NULL,
    @CodigoDocumento NVARCHAR(4) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @PaginaNorm INT = CASE WHEN @Pagina IS NULL OR @Pagina < 1 THEN 1 ELSE @Pagina END;
        DECLARE @TamanoNorm INT = CASE WHEN @TamanoPagina IS NULL OR @TamanoPagina < 1 THEN 20 ELSE @TamanoPagina END;
        DECLARE @Offset INT = (@PaginaNorm - 1) * @TamanoNorm;
        DECLARE @BuscarNorm NVARCHAR(120) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');
        DECLARE @CodigoDocumentoNorm NVARCHAR(4) = NULLIF(UPPER(LTRIM(RTRIM(@CodigoDocumento))), N'');

        CREATE TABLE #Base
        (
            Id INT NOT NULL,
            Tipo NVARCHAR(50) NOT NULL,
            SerieNumero NVARCHAR(30) NOT NULL,
            FechaEmision DATETIME2 NOT NULL,
            Cliente NVARCHAR(200) NOT NULL,
            Total DECIMAL(10,2) NOT NULL,
            Estado NVARCHAR(50) NOT NULL,
            ReservaId INT NOT NULL,
            EsTributario BIT NOT NULL,
            UrlDescargaProveedor NVARCHAR(500) NULL
        );

        INSERT INTO #Base
        (
            Id,
            Tipo,
            SerieNumero,
            FechaEmision,
            Cliente,
            Total,
            Estado,
            ReservaId,
            EsTributario,
            UrlDescargaProveedor
        )
        SELECT
            c.Id,
            CASE
                WHEN c.TipoComprobante = 2 THEN N'Factura'
                WHEN c.TipoComprobante = 1 THEN N'Boleta'
                WHEN c.TipoComprobante = 3 THEN N'Recibo Interno'
                ELSE CONCAT(N'Tipo ', c.TipoComprobante)
            END AS Tipo,
            CONCAT(c.Serie, N'-', c.Numero) AS SerieNumero,
            c.FechaEmision,
            cl.NombresORazonSocial AS Cliente,
            c.Total,
            CASE
                WHEN c.Estado = 1 THEN N'Pendiente'
                WHEN c.Estado = 2 THEN N'Enviado'
                WHEN c.Estado = 3 THEN N'Aceptado'
                WHEN c.Estado = 4 THEN N'Rechazado'
                WHEN c.Estado = 5 THEN N'Anulado'
                ELSE CONCAT(N'Estado ', c.Estado)
            END AS Estado,
            ISNULL(c.ReservaId, 0) AS ReservaId,
            CAST(CASE WHEN c.TipoComprobante IN (1, 2) THEN 1 ELSE 0 END AS BIT) AS EsTributario,
            CASE WHEN c.MensajeRespuestaSunat LIKE N'http%' THEN c.MensajeRespuestaSunat ELSE NULL END AS UrlDescargaProveedor
        FROM dbo.ComprobantesElectronicos c
        INNER JOIN dbo.Clientes cl ON cl.Id = c.ClienteId
        LEFT JOIN dbo.Reservas r ON r.Id = c.ReservaId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        LEFT JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE c.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND
          (
                @CodigoDocumentoNorm IS NULL
                OR
                (
                    CASE
                        WHEN c.TipoComprobante = 2 THEN N'01'
                        WHEN c.TipoComprobante = 1 THEN N'03'
                        WHEN c.TipoComprobante = 3 THEN N'RI'
                        ELSE NULL
                    END
                ) = @CodigoDocumentoNorm
          )
          AND
          (
                @BuscarNorm IS NULL
                OR CONCAT(c.Serie, N'-', c.Numero) LIKE N'%' + @BuscarNorm + N'%'
                OR cl.NombresORazonSocial LIKE N'%' + @BuscarNorm + N'%'
                OR CAST(ISNULL(c.ReservaId, 0) AS NVARCHAR(20)) LIKE N'%' + @BuscarNorm + N'%'
                OR
                (
                    CASE
                        WHEN c.TipoComprobante = 2 THEN N'Factura'
                        WHEN c.TipoComprobante = 1 THEN N'Boleta'
                        WHEN c.TipoComprobante = 3 THEN N'Recibo Interno'
                        ELSE CONCAT(N'Tipo ', c.TipoComprobante)
                    END
                ) LIKE N'%' + @BuscarNorm + N'%'
                OR
                (
                    CASE
                        WHEN c.Estado = 1 THEN N'Pendiente'
                        WHEN c.Estado = 2 THEN N'Enviado'
                        WHEN c.Estado = 3 THEN N'Aceptado'
                        WHEN c.Estado = 4 THEN N'Rechazado'
                        WHEN c.Estado = 5 THEN N'Anulado'
                        ELSE CONCAT(N'Estado ', c.Estado)
                    END
                ) LIKE N'%' + @BuscarNorm + N'%'
          );

        SELECT @TotalRegistros = COUNT(1)
        FROM #Base;

        SELECT
            b.Id,
            b.Tipo,
            b.SerieNumero,
            b.FechaEmision,
            b.Cliente,
            b.Total,
            b.Estado,
            b.ReservaId,
            b.EsTributario,
            b.UrlDescargaProveedor
        FROM #Base b
        ORDER BY b.FechaEmision DESC, b.Id DESC
        OFFSET @Offset ROWS
        FETCH NEXT @TamanoNorm ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
