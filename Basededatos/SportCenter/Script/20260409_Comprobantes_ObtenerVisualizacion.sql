/*
Firma: Codex - 09/04/2026
Descripcion: Crea/actualiza SP para visualizar comprobante (PDF interno o redireccion URL proveedor en tributarios), incluyendo ubigeo de negocio y cliente.
Firma: Codex - 07/05/2026 | Corrige EsTributario usando CodigoSunat por negocio (sin IDs fijos).
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Comprobantes_ObtenerVisualizacion]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            ce.Id,
            ce.NegocioId,
            ce.ReservaId,
            ce.TipoComprobante,
            CASE
                WHEN ce.TipoComprobante = 2 THEN N'01'
                WHEN ce.TipoComprobante = 1 THEN N'03'
                WHEN ce.TipoComprobante = 3 THEN N'RI'
                ELSE N'03'
            END AS CodigoDocumentoComprobante,
            CASE
                WHEN ce.TipoComprobante = 2 THEN N'Factura'
                WHEN ce.TipoComprobante = 1 THEN N'Boleta'
                WHEN ce.TipoComprobante = 3 THEN N'Recibo Interno'
                ELSE CONCAT(N'Tipo ', ce.TipoComprobante)
            END AS TipoDocumentoNombre,
            CAST(CASE WHEN d.CodigoSunat IN (N'01', N'03', N'07', N'08') THEN 1 ELSE 0 END AS BIT) AS EsTributario,
            ce.Serie,
            ce.Numero,
            ce.FechaEmision,
            COALESCE(ms.Simbolo, N'S/') AS MonedaSimbolo,
            ce.SubTotal,
            ce.Igv,
            ce.Total,
            ISNULL(n.PorcentajeIgv, 18) AS PorcentajeIgv,
            n.NombreComercial,
            n.RazonSocial,
            n.DireccionFiscal,
            ndis.Nombre AS NegocioDistrito,
            nprov.Nombre AS NegocioProvincia,
            ndep.Nombre AS NegocioDepartamento,
            CASE
                WHEN ISNULL(n.TipoDocumentoFiscal, N'') <> N'' OR ISNULL(n.NumeroDocumentoFiscal, N'') <> N''
                    THEN CONCAT(ISNULL(n.TipoDocumentoFiscal, N''), CASE WHEN ISNULL(n.NumeroDocumentoFiscal, N'') = N'' THEN N'' ELSE N'-' + n.NumeroDocumentoFiscal END)
                ELSE ISNULL(n.DocumentoFiscal, N'')
            END AS NegocioDocumento,
            c.NombresORazonSocial AS ClienteNombre,
            CONCAT(ISNULL(c.TipoDocumento, N''), CASE WHEN ISNULL(c.NumeroDocumento, N'') = N'' THEN N'' ELSE N'-' + c.NumeroDocumento END) AS ClienteDocumento,
            c.DireccionFiscal AS ClienteDireccion,
            cdis.Nombre AS ClienteDistrito,
            cprov.Nombre AS ClienteProvincia,
            cdep.Nombre AS ClienteDepartamento,
            c.Correo AS ClienteCorreo,
            s.Nombre AS SedeNombre,
            e.Nombre AS EspacioNombre,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            CASE WHEN ce.MensajeRespuestaSunat LIKE N'http%' THEN ce.MensajeRespuestaSunat ELSE NULL END AS UrlDescargaProveedor
        FROM dbo.ComprobantesElectronicos ce
        INNER JOIN dbo.Negocios n ON n.Id = ce.NegocioId
        INNER JOIN dbo.Reservas r ON r.Id = ce.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = ce.ClienteId
        INNER JOIN dbo.NegociosTiposDocumentoComprobante d
            ON d.Id = ce.TipoComprobante
           AND d.NegocioId = ce.NegocioId
        LEFT JOIN dbo.UbigeoDistritos ndis ON ndis.CodigoUbigeo = n.CodigoUbigeo
        LEFT JOIN dbo.UbigeoProvincias nprov ON nprov.CodigoProvincia = ndis.CodigoProvincia
        LEFT JOIN dbo.UbigeoDepartamentos ndep ON ndep.CodigoDepartamento = ndis.CodigoDepartamento
        LEFT JOIN dbo.UbigeoDistritos cdis ON cdis.CodigoUbigeo = c.CodigoUbigeo
        LEFT JOIN dbo.UbigeoProvincias cprov ON cprov.CodigoProvincia = cdis.CodigoProvincia
        LEFT JOIN dbo.UbigeoDepartamentos cdep ON cdep.CodigoDepartamento = cdis.CodigoDepartamento
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Id = m.MonedaSuperId
        WHERE ce.NegocioId = @NegocioId
          AND ce.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
