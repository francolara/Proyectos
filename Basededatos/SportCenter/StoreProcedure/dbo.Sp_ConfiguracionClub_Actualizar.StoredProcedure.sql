USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_ConfiguracionClub_Actualizar por ubigeo fiscal y tipo de documento SUNAT centralizado.
-- Firma: Codex - 06/04/2026 | Se agrega configuracion de politica de confirmacion por pago y porcentaje minimo de adelanto por negocio.
-- Firma: Codex - 09/04/2026 | Se agrega configuracion de emision (CPE/Recibo interno) y porcentaje IGV.
-- Firma: Codex - 13/04/2026 | Se agrega persistencia de LogoUrl para logo del negocio.
-- Firma: Codex - 16/04/2026 | Se agregan flags de reserva: permitir modificar precio y cancelacion automatica por no confirmacion.
-- Firma: Codex - 06/05/2026 | Se agrega persistencia de EnviarComprobanteAutomatico desde configuracion del negocio.
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @MonedaId INT,
    @PoliticaConfirmacionPago TINYINT = 0,
    @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL,
    @EmisionComprobantesElectronicos BIT = 0,
    @EnviarComprobanteAutomatico BIT = 0,
    @EmisionReciboInterno BIT = 0,
    @PorcentajeIgv INT = 18,
    @LogoUrl NVARCHAR(500) = NULL,
    @PermitirModificarPrecioReserva BIT = 0,
    @CancelacionAutomaticaNoConfirmada BIT = 0,
    @MinutosCancelacionNoConfirmada INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @DireccionFiscalNormalizada NVARCHAR(250);
        DECLARE @CodigoUbigeoNormalizado CHAR(6);
        DECLARE @PorcentajeAdelantoNormalizado DECIMAL(5,2);
        DECLARE @LogoUrlNormalizado NVARCHAR(500);

        SET @TipoDocumentoFiscal = NULLIF(UPPER(LTRIM(RTRIM(@TipoDocumentoFiscal))), N'');
        SET @DireccionFiscalNormalizada = NULLIF(LTRIM(RTRIM(@DireccionFiscal)), N'');
        SET @CodigoUbigeoNormalizado = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @PorcentajeAdelantoNormalizado = @PorcentajeAdelantoMinimo;
        SET @LogoUrlNormalizado = NULLIF(LTRIM(RTRIM(@LogoUrl)), N'');

        IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Id = @MonedaId AND Activo = 1)
            RAISERROR('La moneda seleccionada no es valida.', 16, 1);

        IF @TipoDocumentoFiscal IS NULL
            RAISERROR('El tipo de documento SUNAT es obligatorio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = @TipoDocumentoFiscal AND t.Activo = 1)
            RAISERROR('El tipo de documento SUNAT no es valido.', 16, 1);

        IF @DireccionFiscalNormalizada IS NULL
            SET @CodigoUbigeoNormalizado = NULL;

        IF @DireccionFiscalNormalizada IS NOT NULL AND @CodigoUbigeoNormalizado IS NULL
            RAISERROR('Cuando se registra direccion fiscal, el distrito es obligatorio.', 16, 1);

        IF @CodigoUbigeoNormalizado IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeoNormalizado AND Activo = 1)
            RAISERROR('El codigo de ubigeo no existe.', 16, 1);

        IF @PoliticaConfirmacionPago NOT IN (0, 1, 2)
            RAISERROR('La politica de confirmacion no es valida.', 16, 1);

        IF @PoliticaConfirmacionPago = 1
        BEGIN
            IF @PorcentajeAdelantoNormalizado IS NULL OR @PorcentajeAdelantoNormalizado < 1 OR @PorcentajeAdelantoNormalizado > 100
                RAISERROR('Para exigir adelanto, el porcentaje minimo debe ser entero entre 1 y 100.', 16, 1);

            IF @PorcentajeAdelantoNormalizado <> FLOOR(@PorcentajeAdelantoNormalizado)
                RAISERROR('El porcentaje minimo de adelanto no admite decimales.', 16, 1);
        END
        ELSE
        BEGIN
            SET @PorcentajeAdelantoNormalizado = NULL;
        END

        IF @PorcentajeIgv IS NULL OR @PorcentajeIgv < 0 OR @PorcentajeIgv > 100
            RAISERROR('El porcentaje de IGV debe estar entre 0 y 100.', 16, 1);

        IF @CancelacionAutomaticaNoConfirmada = 1
        BEGIN
            IF @MinutosCancelacionNoConfirmada IS NULL OR @MinutosCancelacionNoConfirmada < 5 OR @MinutosCancelacionNoConfirmada > 1440
                RAISERROR('El tiempo de cancelacion automatica debe estar entre 5 y 1440 minutos.', 16, 1);
        END
        ELSE
        BEGIN
            SET @MinutosCancelacionNoConfirmada = NULL;
        END

        IF @EmisionComprobantesElectronicos = 1
           AND NOT EXISTS (
                SELECT 1
                FROM dbo.NegociosTiposDocumentoComprobante ntd
                INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
                WHERE ntd.NegocioId = @NegocioId
                  AND ntd.Activo = 1
                  AND t.Activo = 1
                  AND t.Habilitado = 1
                  AND t.Tributario = 1
            )
            RAISERROR('Debes habilitar al menos un documento tributario en Maestros para activar emision de comprobantes electronicos.', 16, 1);

        IF @EmisionReciboInterno = 1
           AND NOT EXISTS (
                SELECT 1
                FROM dbo.NegociosTiposDocumentoComprobante ntd
                INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
                WHERE ntd.NegocioId = @NegocioId
                  AND ntd.Activo = 1
                  AND t.Activo = 1
                  AND t.Habilitado = 1
                  AND t.Tributario = 0
            )
            RAISERROR('Debes habilitar al menos un documento no tributario en Maestros para activar emision de recibo interno.', 16, 1);

        UPDATE n
        SET
            n.NombreComercial = @NombreComercial,
            n.RazonSocial = NULLIF(@RazonSocial, N''),
            n.TipoDocumentoFiscal = @TipoDocumentoFiscal,
            n.NumeroDocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.DireccionFiscal = @DireccionFiscalNormalizada,
            n.CodigoUbigeo = @CodigoUbigeoNormalizado,
            n.DocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.MonedaId = @MonedaId,
            n.PoliticaConfirmacionPago = @PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo = @PorcentajeAdelantoNormalizado,
            n.EmisionComprobantesElectronicos = @EmisionComprobantesElectronicos,
            n.EnviarComprobanteAutomatico = @EnviarComprobanteAutomatico,
            n.EmisionReciboInterno = @EmisionReciboInterno,
            n.PorcentajeIgv = @PorcentajeIgv,
            n.LogoUrl = @LogoUrlNormalizado,
            n.PermitirModificarPrecioReserva = @PermitirModificarPrecioReserva,
            n.CancelacionAutomaticaNoConfirmada = @CancelacionAutomaticaNoConfirmada,
            n.MinutosCancelacionNoConfirmada = @MinutosCancelacionNoConfirmada
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80);
        SET @EntidadIdAuditoria = CONVERT(NVARCHAR(80), @NegocioId);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
