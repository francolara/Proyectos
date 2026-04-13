USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Comprobantes_Crear para usar codigo SUNAT real del cliente.
-- Firma: Codex - 09/04/2026 | Emision de comprobantes por series configuradas (CPE/Recibo interno), con correlativo automatico por tipo/serie, validaciones tributarias y actualizacion de datos de cliente desde comprobante.
-- Firma: Codex - 11/04/2026 | Soporta emision de NC/ND desde comprobante referencia (Factura/Boleta aceptada SUNAT), con tipo de nota SUNAT.
-- Firma: Codex - 12/04/2026 | Elimina mapeos fijos de TipoComprobante (1/2/3/4/5) y resuelve dinamicamente CodigoSunat/TipoComprobanteId por negocio en NegociosTiposDocumentoComprobante.
-- Firma: Codex - 13/04/2026 | Permite reemision de comprobante principal cuando el comprobante inicial tiene NC activa, sin anular el comprobante inicial.
CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Crear
    @NegocioId INT,
    @ReservaId INT,
    @TipoComprobante INT,
    @CodigoDocumentoComprobante NVARCHAR(4) = NULL,
    @NegocioSerieId INT = NULL,
    @Serie NVARCHAR(4),
    @Numero INT = NULL,
    @FechaEmision DATETIME2,
    @TipoMoneda INT,
    @SubTotal DECIMAL(10,2),
    @Igv DECIMAL(10,2),
    @Total DECIMAL(10,2),
    @Estado INT,
    @ClienteCorreo NVARCHAR(200) = NULL,
    @ClienteTipoDocumento NVARCHAR(20) = NULL,
    @ClienteNumeroDocumento NVARCHAR(20) = NULL,
    @ClienteDireccionFiscal NVARCHAR(250) = NULL,
    @ClienteCodigoUbigeo CHAR(6) = NULL,
    @ComprobanteReferenciaId INT = NULL,
    @TipoNota CHAR(2) = NULL,
    @TipoNotaCodigoSunat NVARCHAR(2) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @ClienteId INT;
        DECLARE @CodigoTipoDocumentoClienteSunat NVARCHAR(2);
        DECLARE @ReservaEstado INT;
        DECLARE @SedeId INT;
        DECLARE @CodigoDoc NVARCHAR(4);
        DECLARE @SerieConfigurada NVARCHAR(4);
        DECLARE @Tributario BIT;
        DECLARE @NumeroGenerado INT;
        DECLARE @TipoDocumentoClienteFinal NVARCHAR(20);
        DECLARE @NumeroDocumentoClienteFinal NVARCHAR(20);
        DECLARE @DireccionFiscalClienteFinal NVARCHAR(250);
        DECLARE @CodigoUbigeoClienteFinal CHAR(6);
        DECLARE @EsNota BIT = 0;
        DECLARE @TipoNotaNorm CHAR(2);
        DECLARE @TipoComprobanteReferenciaCodigo NVARCHAR(4);
        DECLARE @EstadoComprobanteReferencia INT;
        DECLARE @ReservaReferenciaId INT;
        DECLARE @ClienteReferenciaId INT;

        SET @CodigoDoc = UPPER(LTRIM(RTRIM(ISNULL(@CodigoDocumentoComprobante, N''))));
        SET @TipoNotaNorm = UPPER(LTRIM(RTRIM(ISNULL(@TipoNota, ''))));
        SET @TipoNotaCodigoSunat = NULLIF(UPPER(LTRIM(RTRIM(@TipoNotaCodigoSunat))), N'');

        IF @CodigoDoc = N''
        BEGIN
            SELECT TOP (1)
                @CodigoDoc = ntd.CodigoSunat
            FROM dbo.NegociosTiposDocumentoComprobante ntd
            WHERE ntd.Id = @TipoComprobante
              AND ntd.NegocioId = @NegocioId
              AND ntd.Activo = 1;
        END

        IF @CodigoDoc = N''
            RAISERROR('No se pudo determinar el tipo de documento del comprobante.', 16, 1);

        IF @CodigoDoc IN (N'07', N'08')
            SET @EsNota = 1;

        IF @TipoNotaNorm IN ('NC', 'ND', '07', '08')
            SET @EsNota = 1;

        IF @EsNota = 1
        BEGIN
            IF @TipoNotaNorm = ''
                SET @TipoNotaNorm = CASE WHEN @CodigoDoc = N'08' THEN '08' ELSE '07' END;

            IF @TipoNotaNorm = 'NC' SET @TipoNotaNorm = '07';
            IF @TipoNotaNorm = 'ND' SET @TipoNotaNorm = '08';

            IF @TipoNotaNorm NOT IN ('07', '08')
                RAISERROR('El tipo de nota debe ser 07 o 08.', 16, 1);

            IF @CodigoDoc = N'07' AND @TipoNotaNorm <> '07'
                RAISERROR('El documento 07 requiere tipo de nota 07.', 16, 1);

            IF @CodigoDoc = N'08' AND @TipoNotaNorm <> '08'
                RAISERROR('El documento 08 requiere tipo de nota 08.', 16, 1);

            IF @ComprobanteReferenciaId IS NULL OR @ComprobanteReferenciaId <= 0
                RAISERROR('Para NC/ND se requiere comprobante de referencia.', 16, 1);

            IF @TipoNotaCodigoSunat IS NULL
                RAISERROR('Selecciona el tipo de nota SUNAT.', 16, 1);

            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.TiposNotaComprobanteSunat t
                WHERE t.TipoNota = @TipoNotaNorm
                  AND t.CodigoSunat = @TipoNotaCodigoSunat
                  AND t.Activo = 1
            )
                RAISERROR('El tipo de nota SUNAT no es valido.', 16, 1);

            SELECT
                @TipoComprobanteReferenciaCodigo = ntdRef.CodigoSunat,
                @EstadoComprobanteReferencia = ce.Estado,
                @ReservaReferenciaId = ce.ReservaId,
                @ClienteReferenciaId = ce.ClienteId
            FROM dbo.ComprobantesElectronicos ce
            LEFT JOIN dbo.NegociosTiposDocumentoComprobante ntdRef ON ntdRef.Id = ce.TipoComprobante
            WHERE ce.Id = @ComprobanteReferenciaId
              AND ce.NegocioId = @NegocioId;

            IF @TipoComprobanteReferenciaCodigo IS NULL
                RAISERROR('No se encontro el comprobante de referencia.', 16, 1);

            IF @EstadoComprobanteReferencia <> 3
                RAISERROR('Solo se permite generar NC/ND cuando el comprobante de referencia esta aceptado en SUNAT.', 16, 1);

            IF @TipoComprobanteReferenciaCodigo NOT IN (N'01', N'03')
                RAISERROR('Solo Factura o Boleta pueden ser documento de referencia para NC/ND.', 16, 1);

            SET @ReservaId = @ReservaReferenciaId;
            SET @ClienteId = @ClienteReferenciaId;
        END

        SELECT
            @ClienteId = ISNULL(@ClienteId, r.ClienteId),
            @CodigoTipoDocumentoClienteSunat = c.TipoDocumento,
            @ReservaEstado = r.Estado,
            @SedeId = e.SedeId
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @ClienteId IS NULL
            RAISERROR('No se encontro la reserva para generar el comprobante.', 16, 1);

        IF @EsNota = 0 AND @ReservaEstado <> 4
            RAISERROR('Solo se pueden emitir comprobantes sobre reservas pagadas.', 16, 1);

        IF @EsNota = 0
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.ComprobantesElectronicos ce
                WHERE ce.NegocioId = @NegocioId
                  AND ce.ReservaId = @ReservaId
                  AND ce.ComprobanteReferenciaId IS NULL
                  AND ce.Estado <> 5
                  AND NOT EXISTS
                  (
                      SELECT 1
                      FROM dbo.ComprobantesElectronicos nc
                      INNER JOIN dbo.NegociosTiposDocumentoComprobante ntdNc ON ntdNc.Id = nc.TipoComprobante
                      WHERE nc.NegocioId = ce.NegocioId
                        AND nc.ComprobanteReferenciaId = ce.Id
                        AND nc.Estado <> 5
                        AND ntdNc.CodigoSunat = N'07'
                  )
            )
                RAISERROR('La reserva ya tiene un comprobante emitido. No se permite duplicar comprobantes.', 16, 1);
        END

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.TiposDocumentoIdentidadSunat t
            WHERE t.CodigoSunat = @CodigoTipoDocumentoClienteSunat
              AND t.Activo = 1
        )
            RAISERROR('El cliente no tiene un tipo de documento SUNAT valido.', 16, 1);

        SELECT TOP (1)
            @TipoComprobante = ntd.Id,
            @Tributario = t.Tributario
        FROM dbo.NegociosTiposDocumentoComprobante ntd
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
        WHERE ntd.NegocioId = @NegocioId
          AND ntd.CodigoSunat = @CodigoDoc
          AND ntd.Activo = 1
          AND t.Activo = 1
          AND t.Habilitado = 1;

        IF @TipoComprobante IS NULL OR @Tributario IS NULL
            RAISERROR('El tipo de documento no esta habilitado para este negocio.', 16, 1);

        IF @NegocioSerieId IS NOT NULL
        BEGIN
            SELECT TOP (1)
                @SerieConfigurada = ns.Serie
            FROM dbo.NegociosSeriesDocumentoComprobante ns
            WHERE ns.Id = @NegocioSerieId
              AND ns.NegocioId = @NegocioId
              AND ns.CodigoSunat = @CodigoDoc
              AND ns.Activo = 1;

            IF @SerieConfigurada IS NULL
                RAISERROR('La serie seleccionada no pertenece al documento elegido.', 16, 1);

            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.SedesSeriesDocumentoComprobante ss
                WHERE ss.SedeId = @SedeId
                  AND ss.CodigoSunat = @CodigoDoc
                  AND ss.NegocioSerieId = @NegocioSerieId
                  AND ss.Activo = 1
            )
                RAISERROR('La serie seleccionada no esta habilitada para la sede de la reserva.', 16, 1);
        END
        ELSE
        BEGIN
            SELECT TOP (1)
                @SerieConfigurada = ns.Serie
            FROM dbo.SedesSeriesDocumentoComprobante ss
            INNER JOIN dbo.NegociosSeriesDocumentoComprobante ns ON ns.Id = ss.NegocioSerieId
            WHERE ss.SedeId = @SedeId
              AND ss.CodigoSunat = @CodigoDoc
              AND ss.Activo = 1
              AND ns.NegocioId = @NegocioId
              AND ns.Activo = 1
            ORDER BY ns.Serie;

            IF @SerieConfigurada IS NULL
                RAISERROR('No existe serie habilitada para ese documento en la sede de la reserva.', 16, 1);

            SET @Serie = @SerieConfigurada;
        END

        SET @ClienteTipoDocumento = NULLIF(LTRIM(RTRIM(@ClienteTipoDocumento)), N'');
        SET @ClienteNumeroDocumento = NULLIF(LTRIM(RTRIM(@ClienteNumeroDocumento)), N'');
        SET @ClienteDireccionFiscal = NULLIF(LTRIM(RTRIM(@ClienteDireccionFiscal)), N'');
        SET @ClienteCodigoUbigeo = NULLIF(LTRIM(RTRIM(@ClienteCodigoUbigeo)), '');
        SET @ClienteCorreo = NULLIF(LTRIM(RTRIM(@ClienteCorreo)), N'');

        SET @TipoDocumentoClienteFinal = ISNULL(@ClienteTipoDocumento, @CodigoTipoDocumentoClienteSunat);
        SET @NumeroDocumentoClienteFinal = @ClienteNumeroDocumento;
        SET @DireccionFiscalClienteFinal = @ClienteDireccionFiscal;
        SET @CodigoUbigeoClienteFinal = @ClienteCodigoUbigeo;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.TiposDocumentoIdentidadSunat t
            WHERE t.CodigoSunat = @TipoDocumentoClienteFinal
              AND t.Activo = 1
        )
            RAISERROR('El tipo de documento del cliente no es valido.', 16, 1);

        IF @CodigoDoc = N'01' AND @TipoDocumentoClienteFinal <> N'6'
            RAISERROR('Para Factura el cliente debe tener documento RUC (6).', 16, 1);

        IF @CodigoDoc = N'03'
        BEGIN
            IF @Total > 700
                RAISERROR('La boleta no puede exceder S/ 700.00.', 16, 1);

            IF @TipoDocumentoClienteFinal NOT IN (N'0', N'1')
                RAISERROR('Para Boleta el cliente debe tener tipo de documento 0 o 1.', 16, 1);
        END

        IF @TipoDocumentoClienteFinal = N'0'
            SET @NumeroDocumentoClienteFinal = NULL;
        ELSE IF @NumeroDocumentoClienteFinal IS NULL
            RAISERROR('Ingresa el numero de documento del cliente.', 16, 1);

        IF @DireccionFiscalClienteFinal IS NULL
            SET @CodigoUbigeoClienteFinal = NULL;

        IF @DireccionFiscalClienteFinal IS NOT NULL AND @CodigoUbigeoClienteFinal IS NULL
            RAISERROR('Cuando se ingresa direccion, el ubigeo es obligatorio.', 16, 1);

        IF @CodigoUbigeoClienteFinal IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeoClienteFinal AND Activo = 1)
            RAISERROR('El codigo de ubigeo del cliente no existe.', 16, 1);

        BEGIN TRANSACTION;

        IF @CodigoDoc = N'RI'
        BEGIN
            SET @Igv = 0;
            SET @SubTotal = @Total;
        END

        UPDATE dbo.Clientes
        SET
            Correo = @ClienteCorreo,
            TipoDocumento = @TipoDocumentoClienteFinal,
            NumeroDocumento = ISNULL(@NumeroDocumentoClienteFinal, N''),
            DireccionFiscal = @DireccionFiscalClienteFinal,
            CodigoUbigeo = @CodigoUbigeoClienteFinal,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ClienteId
          AND NegocioId = @NegocioId;

        SELECT @NumeroGenerado = ISNULL(MAX(ce.Numero), 0) + 1
        FROM dbo.ComprobantesElectronicos ce WITH (UPDLOCK, HOLDLOCK)
        WHERE ce.NegocioId = @NegocioId
          AND ce.TipoComprobante = @TipoComprobante
          AND ce.Serie = @Serie;

        INSERT INTO dbo.ComprobantesElectronicos
        (
            NegocioId, ReservaId, ClienteId, TipoComprobante, Serie, Numero,
            FechaEmision, TipoMoneda, CodigoTipoOperacionSunat, CodigoTipoDocumentoClienteSunat,
            SubTotal, Igv, Total, Estado, ComprobanteReferenciaId, TipoNota, TipoNotaCodigoSunat,
            FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @ReservaId, @ClienteId, @TipoComprobante, @Serie, @NumeroGenerado,
            @FechaEmision, @TipoMoneda, N'0101', @TipoDocumentoClienteFinal,
            @SubTotal, @Igv, @Total, @Estado, @ComprobanteReferenciaId,
            CASE WHEN @EsNota = 1 THEN @TipoNotaNorm ELSE NULL END,
            CASE WHEN @EsNota = 1 THEN @TipoNotaCodigoSunat ELSE NULL END,
            SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();
        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'COMPROBANTES', @Accion = N'CREATE', @Entidad = N'ComprobanteElectronico', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
