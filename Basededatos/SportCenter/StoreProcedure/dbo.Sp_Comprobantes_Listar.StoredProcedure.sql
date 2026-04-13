USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- Firma: Codex - 09/04/2026 | Lista comprobantes con filtro y paginacion backend (20x20), manteniendo numero de reserva para UI y total de registros para navegacion.
-- Firma: Codex - 10/04/2026 | Agrega filtro por tipo de documento del comprobante (codigo SUNAT) para listado por negocio.
-- Firma: Codex - 11/04/2026 | Incluye codigo/estado numerico y soporta etiquetas de Nota de Credito y Nota de Debito.
-- Firma: Codex - 11/04/2026 | Incluye columna Referencia y bandera TieneNotasRelacionadas para reglas UI de NC/ND y Anular.
-- Firma: Codex - 12/04/2026 | Elimina mapeos rigidos por Id de tipo comprobante y usa relacion NegociosTiposDocumentoComprobante + TiposDocumentoComprobanteSuperMaestro para tipo/codigo/referencia/filtros en entorno multi-negocio.
-- Firma: Codex - 12/04/2026 | En columna Referencia usa abreviatura del documento desde TiposDocumentoComprobanteSuperMaestro.Abreviatura.
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
            EstadoCodigo INT NOT NULL,
            CodigoDocumentoComprobante NVARCHAR(4) NOT NULL,
            Referencia NVARCHAR(500) NOT NULL,
            TieneNotasRelacionadas BIT NOT NULL,
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
            EstadoCodigo,
            CodigoDocumentoComprobante,
            Referencia,
            TieneNotasRelacionadas,
            ReservaId,
            EsTributario,
            UrlDescargaProveedor
        )
        SELECT
            c.Id,
            COALESCE(tdsm.Nombre, CONCAT(N'Tipo ', ntd.CodigoSunat)) AS Tipo,
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
            c.Estado AS EstadoCodigo,
            COALESCE(ntd.CodigoSunat, N'') AS CodigoDocumentoComprobante,
            CASE
                WHEN ntd.CodigoSunat IN (N'07', N'08') AND cref.Id IS NOT NULL
                    THEN CONCAT(
                        COALESCE(tdsmRef.Abreviatura, tdsmRef.Nombre, N'Comp.'),
                        N' ',
                        cref.Serie,
                        N'-',
                        FORMAT(cref.Numero, '00000000'))
                WHEN notasRelacionadas.RefTexto IS NOT NULL
                    THEN notasRelacionadas.RefTexto
                ELSE N'-'
            END AS Referencia,
            CAST(CASE WHEN notasRelacionadas.TieneNotas = 1 THEN 1 ELSE 0 END AS BIT) AS TieneNotasRelacionadas,
            ISNULL(c.ReservaId, 0) AS ReservaId,
            CAST(COALESCE(tdsm.Tributario, 0) AS BIT) AS EsTributario,
            CASE WHEN c.MensajeRespuestaSunat LIKE N'http%' THEN c.MensajeRespuestaSunat ELSE NULL END AS UrlDescargaProveedor
        FROM dbo.ComprobantesElectronicos c
        INNER JOIN dbo.Clientes cl ON cl.Id = c.ClienteId
        LEFT JOIN dbo.NegociosTiposDocumentoComprobante ntd ON ntd.Id = c.TipoComprobante
        LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = ntd.CodigoSunat
        LEFT JOIN dbo.ComprobantesElectronicos cref ON cref.Id = c.ComprobanteReferenciaId
        LEFT JOIN dbo.NegociosTiposDocumentoComprobante ntdRef ON ntdRef.Id = cref.TipoComprobante
        LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsmRef ON tdsmRef.CodigoSunat = ntdRef.CodigoSunat
        LEFT JOIN dbo.Reservas r ON r.Id = c.ReservaId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        LEFT JOIN dbo.Sedes s ON s.Id = e.SedeId
        OUTER APPLY
        (
            SELECT
                CAST(CASE WHEN COUNT(1) > 0 THEN 1 ELSE 0 END AS INT) AS TieneNotas,
                STRING_AGG(
                    CONCAT(
                        CASE
                            WHEN ntdn.CodigoSunat = N'07' THEN N'NC'
                            WHEN ntdn.CodigoSunat = N'08' THEN N'ND'
                            ELSE N'N'
                        END,
                        N' ',
                        n.Serie,
                        N'-',
                        FORMAT(n.Numero, '00000000')
                    ),
                    N' | '
                ) AS RefTexto
            FROM dbo.ComprobantesElectronicos n
            LEFT JOIN dbo.NegociosTiposDocumentoComprobante ntdn ON ntdn.Id = n.TipoComprobante
            WHERE n.ComprobanteReferenciaId = c.Id
              AND n.NegocioId = c.NegocioId
              AND n.Estado <> 5
              AND ntdn.CodigoSunat IN (N'07', N'08')
        ) AS notasRelacionadas
        WHERE c.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND
          (
                @CodigoDocumentoNorm IS NULL
                OR ntd.CodigoSunat = @CodigoDocumentoNorm
          )
          AND
          (
                @BuscarNorm IS NULL
                OR CONCAT(c.Serie, N'-', c.Numero) LIKE N'%' + @BuscarNorm + N'%'
                OR cl.NombresORazonSocial LIKE N'%' + @BuscarNorm + N'%'
                OR CAST(ISNULL(c.ReservaId, 0) AS NVARCHAR(20)) LIKE N'%' + @BuscarNorm + N'%'
                OR
                COALESCE(tdsm.Nombre, CONCAT(N'Tipo ', c.TipoComprobante)) LIKE N'%' + @BuscarNorm + N'%'
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
            b.EstadoCodigo,
            b.CodigoDocumentoComprobante,
            b.Referencia,
            b.TieneNotasRelacionadas,
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
