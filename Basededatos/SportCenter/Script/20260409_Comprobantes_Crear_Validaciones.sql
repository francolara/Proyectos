/*
Firma: Codex - 09/04/2026
Descripcion: Ajusta emision de comprobantes con correlativo automatico por tipo/serie, validaciones boleta/factura y actualizacion de datos de cliente desde el registro de comprobante.
*/
USE [DbSportCenter]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

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

        SELECT
            @ClienteId = r.ClienteId,
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

        IF @ReservaEstado <> 4
            RAISERROR('Solo se pueden emitir comprobantes sobre reservas pagadas.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.ComprobantesElectronicos ce
            WHERE ce.NegocioId = @NegocioId
              AND ce.ReservaId = @ReservaId
              AND ce.Estado <> 5
        )
            RAISERROR('La reserva ya tiene un comprobante emitido. No se permite duplicar comprobantes.', 16, 1);

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.TiposDocumentoIdentidadSunat t
            WHERE t.CodigoSunat = @CodigoTipoDocumentoClienteSunat
              AND t.Activo = 1
        )
            RAISERROR('El cliente no tiene un tipo de documento SUNAT valido.', 16, 1);

        SET @CodigoDoc = UPPER(LTRIM(RTRIM(ISNULL(@CodigoDocumentoComprobante, N''))));
        IF @CodigoDoc = N''
        BEGIN
            SET @CodigoDoc =
                CASE
                    WHEN @TipoComprobante = 2 THEN N'01'
                    WHEN @TipoComprobante = 1 THEN N'03'
                    WHEN @TipoComprobante = 3 THEN N'RI'
                    ELSE N'03'
                END;
        END

        SELECT TOP (1)
            @Tributario = t.Tributario
        FROM dbo.NegociosTiposDocumentoComprobante ntd
        INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
        WHERE ntd.NegocioId = @NegocioId
          AND ntd.CodigoSunat = @CodigoDoc
          AND ntd.Activo = 1
          AND t.Activo = 1
          AND t.Habilitado = 1;

        IF @Tributario IS NULL
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

        SET @TipoComprobante =
            CASE
                WHEN @CodigoDoc = N'01' THEN 2
                WHEN @CodigoDoc = N'03' THEN 1
                WHEN @CodigoDoc = N'RI' THEN 3
                ELSE @TipoComprobante
            END;

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
            SubTotal, Igv, Total, Estado, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @ReservaId, @ClienteId, @TipoComprobante, @Serie, @NumeroGenerado,
            @FechaEmision, @TipoMoneda, N'0101', @TipoDocumentoClienteFinal,
            @SubTotal, @Igv, @Total, @Estado, SYSUTCDATETIME(), @Usuario
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
