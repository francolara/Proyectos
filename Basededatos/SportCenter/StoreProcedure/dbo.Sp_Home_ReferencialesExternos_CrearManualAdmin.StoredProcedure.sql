USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   12/05/2026
-- Description:   Crea manualmente un referencial externo para Home desde panel superadmin.
-- Firma:         Codex - 12/05/2026 | Alta manual de referencial externo con ubigeo, contacto y coordenadas. GooglePlaceId queda NULL para identificar origen manual.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialesExternos_CrearManualAdmin]
    @NombreComplejo NVARCHAR(180),
    @TipoDeporteSuperId INT,
    @CodigoUbigeo CHAR(6),
    @Direccion NVARCHAR(250) = NULL,
    @TelefonoContacto NVARCHAR(40) = NULL,
    @CorreoContacto NVARCHAR(200) = NULL,
    @LatitudReferencia DECIMAL(10,7) = NULL,
    @LongitudReferencia DECIMAL(10,7) = NULL,
    @GoogleMapsUrl NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @NombreComplejo = NULLIF(LTRIM(RTRIM(@NombreComplejo)), '');
        SET @CodigoUbigeo = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @Direccion = NULLIF(LTRIM(RTRIM(@Direccion)), '');
        SET @TelefonoContacto = NULLIF(LTRIM(RTRIM(@TelefonoContacto)), '');
        SET @CorreoContacto = NULLIF(LTRIM(RTRIM(@CorreoContacto)), '');
        SET @GoogleMapsUrl = NULLIF(LTRIM(RTRIM(@GoogleMapsUrl)), '');
        SET @Usuario = NULLIF(LTRIM(RTRIM(@Usuario)), '');

        IF @NombreComplejo IS NULL
            RAISERROR('El nombre del complejo es obligatorio.', 16, 1);

        IF @TipoDeporteSuperId IS NULL OR @TipoDeporteSuperId <= 0
            RAISERROR('Tipo de deporte invalido.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporteSuperMaestro WHERE Id = @TipoDeporteSuperId)
            RAISERROR('El tipo de deporte no existe.', 16, 1);

        IF @CodigoUbigeo IS NULL OR LEN(@CodigoUbigeo) <> 6
            RAISERROR('Codigo ubigeo invalido.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeo)
            RAISERROR('El codigo ubigeo no existe.', 16, 1);

        IF @LatitudReferencia IS NULL OR @LongitudReferencia IS NULL
            RAISERROR('Debes registrar latitud y longitud.', 16, 1);

        IF @LatitudReferencia < -90 OR @LatitudReferencia > 90
            RAISERROR('Latitud fuera de rango.', 16, 1);

        IF @LongitudReferencia < -180 OR @LongitudReferencia > 180
            RAISERROR('Longitud fuera de rango.', 16, 1);

        INSERT INTO dbo.HomeEspaciosReferencialesExternos
        (
            GooglePlaceId,
            NombreComplejo,
            NombreEspacio,
            CodigoReferencia,
            CodigoUbigeo,
            TipoDeporteSuperId,
            Direccion,
            Referencia,
            TelefonoContacto,
            CorreoContacto,
            WhatsappContacto,
            PermiteChatWhatsapp,
            TarifaReferencial,
            TieneIluminacion,
            Techada,
            GoogleMapsUrl,
            LatitudReferencia,
            LongitudReferencia,
            FotoPrincipalUrl,
            FotosUrlsCsv,
            Activo,
            FechaCreacion,
            UsuarioCreacion
        )
        VALUES
        (
            NULL,
            @NombreComplejo,
            NULL,
            NULL,
            @CodigoUbigeo,
            @TipoDeporteSuperId,
            @Direccion,
            NULL,
            @TelefonoContacto,
            @CorreoContacto,
            NULL,
            0,
            NULL,
            0,
            0,
            @GoogleMapsUrl,
            @LatitudReferencia,
            @LongitudReferencia,
            NULL,
            NULL,
            1,
            SYSUTCDATETIME(),
            COALESCE(@Usuario, N'owner-platform')
        );

        SELECT CAST(SCOPE_IDENTITY() AS INT) AS Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
