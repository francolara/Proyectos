USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/05/2026
-- Description:   Valida checklist minimo obligatorio de onboarding por negocio.
-- =============================================
-- Firma: Codex - 26/05/2026 | Se crea SP de validacion integral para configuracion, maestros, sedes y espacios; se exige en sede notificaciones activas, correo y WhatsApp via dbo.SedeConfiguracionNotificacion.
-- Firma: Codex - 25/05/2026 | Se amplia requisito de Configuracion inicial (razon social, documento, direccion, ubigeo, IGV y reglas de reserva) y se quita documento/serie como requisito obligatorio de Maestros.
CREATE OR ALTER PROCEDURE [dbo].[Sp_OnboardingChecklist_Validar]
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ConfigNombreComercialOk BIT = 0;
        DECLARE @ConfigTipoDocumentoOk BIT = 0;
        DECLARE @ConfigMonedaOk BIT = 0;
        DECLARE @ConfigCpeCondicionesOk BIT = 1;

        DECLARE @MaestroTipoDeporteOk BIT = 0;
        DECLARE @MaestroTipoSueloOk BIT = 0;
        DECLARE @MaestroFormaPagoOk BIT = 0;
        DECLARE @MaestroMonedaOk BIT = 0;
        DECLARE @MaestroTipoDocumentoOk BIT = 0;
        DECLARE @MaestroSerieDocumentoOk BIT = 0;

        DECLARE @SedeMinimaOk BIT = 0;
        DECLARE @EspacioMinimoOk BIT = 0;

        SELECT
            @ConfigNombreComercialOk = CASE WHEN LTRIM(RTRIM(COALESCE(n.NombreComercial, N''))) <> N'' THEN 1 ELSE 0 END,
            @ConfigTipoDocumentoOk = CASE WHEN LTRIM(RTRIM(COALESCE(n.TipoDocumentoFiscal, N''))) <> N'' THEN 1 ELSE 0 END,
            @ConfigMonedaOk = CASE WHEN n.MonedaId IS NOT NULL THEN 1 ELSE 0 END,
            @ConfigCpeCondicionesOk =
                CASE
                    WHEN LTRIM(RTRIM(COALESCE(n.RazonSocial, N''))) <> N''
                         AND LTRIM(RTRIM(COALESCE(n.TipoDocumentoFiscal, N''))) <> N''
                         AND LTRIM(RTRIM(COALESCE(n.NumeroDocumentoFiscal, N''))) <> N''
                         AND LTRIM(RTRIM(COALESCE(n.DireccionFiscal, N''))) <> N''
                         AND n.CodigoUbigeo IS NOT NULL
                         AND LEN(LTRIM(RTRIM(COALESCE(n.CodigoUbigeo, N'')))) = 6
                         AND COALESCE(n.PorcentajeIgv, 0) > 0
                    THEN 1
                    ELSE 0
                END
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        SELECT @MaestroTipoDeporteOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.TiposDeporte td WHERE td.NegocioId = @NegocioId AND td.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @MaestroTipoSueloOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.TiposSuelo ts WHERE ts.NegocioId = @NegocioId AND ts.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @MaestroFormaPagoOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.FormasPago fp WHERE fp.NegocioId = @NegocioId AND fp.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @MaestroMonedaOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.Monedas m WHERE m.NegocioId = @NegocioId AND m.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @MaestroTipoDocumentoOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.NegociosTiposDocumentoComprobante tdc WHERE tdc.NegocioId = @NegocioId AND tdc.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @MaestroSerieDocumentoOk = CASE WHEN EXISTS (
            SELECT 1 FROM dbo.NegociosSeriesDocumentoComprobante sdc WHERE sdc.NegocioId = @NegocioId AND sdc.Activo = 1
        ) THEN 1 ELSE 0 END;

        SELECT @SedeMinimaOk = CASE WHEN EXISTS (
            SELECT 1
            FROM dbo.Sedes s
            INNER JOIN dbo.SedeConfiguracionNotificacion scn
                ON scn.SedeId = s.Id
            WHERE s.NegocioId = @NegocioId
              AND s.Activo = 1
              AND LTRIM(RTRIM(COALESCE(s.Nombre, N''))) <> N''
              AND LTRIM(RTRIM(COALESCE(s.Direccion, N''))) <> N''
              AND s.CodigoUbigeo IS NOT NULL
              AND LEN(LTRIM(RTRIM(COALESCE(s.CodigoUbigeo, N'')))) = 6
              AND COALESCE(scn.NotificacionesActivas, 0) = 1
              AND LTRIM(RTRIM(COALESCE(scn.CorreoNotificacion, N''))) <> N''
              AND LTRIM(RTRIM(COALESCE(scn.WhatsappContacto, N''))) <> N''
              AND EXISTS (
                    SELECT 1
                    FROM dbo.SedeHorarioAtencion sh
                    WHERE sh.SedeId = s.Id
                      AND sh.HoraApertura IS NOT NULL
                      AND sh.HoraCierre IS NOT NULL
                )
              AND EXISTS (
                    SELECT 1
                    FROM dbo.SedeServicios ss
                    WHERE ss.SedeId = s.Id
                )
        ) THEN 1 ELSE 0 END;

        SELECT @EspacioMinimoOk = CASE WHEN EXISTS (
            SELECT 1
            FROM dbo.EspaciosDeportivos e
            INNER JOIN dbo.Sedes s
                ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND s.Activo = 1
              AND e.Estado = 1
              AND LTRIM(RTRIM(COALESCE(e.Codigo, N''))) <> N''
              AND LTRIM(RTRIM(COALESCE(e.Nombre, N''))) <> N''
              AND e.TipoDeporteId > 0
              AND e.TipoSueloId > 0
              AND EXISTS (
                    SELECT 1
                    FROM dbo.Tarifas t
                    WHERE t.EspacioDeportivoId = e.Id
                      AND t.Activa = 1
                      AND t.Precio > 0
                      AND t.HoraFin > t.HoraInicio
                )
        ) THEN 1 ELSE 0 END;

        SELECT
            @NegocioId AS NegocioId,
            @ConfigNombreComercialOk AS ConfigNombreComercialOk,
            @ConfigTipoDocumentoOk AS ConfigTipoDocumentoOk,
            @ConfigMonedaOk AS ConfigMonedaOk,
            @ConfigCpeCondicionesOk AS ConfigCpeCondicionesOk,
            @MaestroTipoDeporteOk AS MaestroTipoDeporteOk,
            @MaestroTipoSueloOk AS MaestroTipoSueloOk,
            @MaestroFormaPagoOk AS MaestroFormaPagoOk,
            @MaestroMonedaOk AS MaestroMonedaOk,
            @MaestroTipoDocumentoOk AS MaestroTipoDocumentoOk,
            @MaestroSerieDocumentoOk AS MaestroSerieDocumentoOk,
            @SedeMinimaOk AS SedeMinimaOk,
            @EspacioMinimoOk AS EspacioMinimoOk,
            CAST(
                CASE WHEN
                    @ConfigNombreComercialOk = 1
                    AND @ConfigTipoDocumentoOk = 1
                    AND @ConfigMonedaOk = 1
                    AND @ConfigCpeCondicionesOk = 1
                    AND @MaestroTipoDeporteOk = 1
                    AND @MaestroTipoSueloOk = 1
                    AND @MaestroFormaPagoOk = 1
                    AND @MaestroMonedaOk = 1
                    AND @SedeMinimaOk = 1
                    AND @EspacioMinimoOk = 1
                THEN 1 ELSE 0 END
            AS BIT) AS ChecklistCompleto;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
