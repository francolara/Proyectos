USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Inserta o actualiza registros de HomeEspaciosReferencialesExternos usando GooglePlaceId como llave de sincronizacion.
-- Firma: Codex - 27/04/2026 | Amplia upsert para persistir TelefonoContacto y coordenadas (LatitudReferencia/LongitudReferencia) capturadas desde Google.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialExterno_UpsertDesdeGoogle]
    @GooglePlaceId NVARCHAR(200),
    @NombreComplejo NVARCHAR(180),
    @NombreEspacio NVARCHAR(150) = NULL,
    @CodigoReferencia NVARCHAR(50) = NULL,
    @CodigoUbigeo CHAR(6),
    @TipoDeporteSuperId INT,
    @Direccion NVARCHAR(250) = NULL,
    @Referencia NVARCHAR(1000) = NULL,
    @TelefonoContacto NVARCHAR(40) = NULL,
    @CorreoContacto NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(30) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @TarifaReferencial DECIMAL(10,2) = NULL,
    @TieneIluminacion BIT = 0,
    @Techada BIT = 0,
    @GoogleMapsUrl NVARCHAR(500) = NULL,
    @FotoPrincipalUrl NVARCHAR(500) = NULL,
    @FotosUrlsCsv NVARCHAR(MAX) = NULL,
    @LatitudReferencia DECIMAL(10,7) = NULL,
    @LongitudReferencia DECIMAL(10,7) = NULL,
    @Activo BIT = 1,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @GooglePlaceId = NULLIF(LTRIM(RTRIM(@GooglePlaceId)), '');
        SET @NombreComplejo = LTRIM(RTRIM(@NombreComplejo));
        SET @NombreEspacio = NULLIF(LTRIM(RTRIM(@NombreEspacio)), '');
        SET @CodigoReferencia = NULLIF(LTRIM(RTRIM(@CodigoReferencia)), '');
        SET @Direccion = NULLIF(LTRIM(RTRIM(@Direccion)), '');
        SET @Referencia = NULLIF(LTRIM(RTRIM(@Referencia)), '');
        SET @TelefonoContacto = NULLIF(LTRIM(RTRIM(@TelefonoContacto)), '');
        SET @CorreoContacto = NULLIF(LTRIM(RTRIM(@CorreoContacto)), '');
        SET @WhatsappContacto = NULLIF(LTRIM(RTRIM(@WhatsappContacto)), '');
        SET @GoogleMapsUrl = NULLIF(LTRIM(RTRIM(@GoogleMapsUrl)), '');
        SET @FotoPrincipalUrl = NULLIF(LTRIM(RTRIM(@FotoPrincipalUrl)), '');
        SET @FotosUrlsCsv = NULLIF(LTRIM(RTRIM(@FotosUrlsCsv)), '');
        SET @Usuario = COALESCE(NULLIF(LTRIM(RTRIM(@Usuario)), ''), 'sync-google');

        IF @GooglePlaceId IS NULL
            RAISERROR('GooglePlaceId es obligatorio para sincronizar referenciales externos.', 16, 1);

        IF @NombreComplejo = ''
            RAISERROR('NombreComplejo es obligatorio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeo)
            RAISERROR('CodigoUbigeo no existe en maestro UBIGEO.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDeporteSuperMaestro WHERE Id = @TipoDeporteSuperId)
            RAISERROR('TipoDeporteSuperId no existe en supermaestro.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.HomeEspaciosReferencialesExternos WHERE GooglePlaceId = @GooglePlaceId)
        BEGIN
            UPDATE dbo.HomeEspaciosReferencialesExternos
               SET NombreComplejo = @NombreComplejo,
                   NombreEspacio = @NombreEspacio,
                   CodigoReferencia = @CodigoReferencia,
                   CodigoUbigeo = @CodigoUbigeo,
                   TipoDeporteSuperId = @TipoDeporteSuperId,
                   Direccion = @Direccion,
                   Referencia = @Referencia,
                   TelefonoContacto = @TelefonoContacto,
                   CorreoContacto = @CorreoContacto,
                   WhatsappContacto = @WhatsappContacto,
                   PermiteChatWhatsapp = COALESCE(@PermiteChatWhatsapp, 0),
                   TarifaReferencial = @TarifaReferencial,
                   TieneIluminacion = COALESCE(@TieneIluminacion, 0),
                   Techada = COALESCE(@Techada, 0),
                   GoogleMapsUrl = @GoogleMapsUrl,
                   LatitudReferencia = @LatitudReferencia,
                   LongitudReferencia = @LongitudReferencia,
                   FotoPrincipalUrl = @FotoPrincipalUrl,
                   FotosUrlsCsv = @FotosUrlsCsv,
                   Activo = COALESCE(@Activo, 1),
                   FechaActualizacion = SYSUTCDATETIME(),
                   UsuarioActualizacion = @Usuario
             WHERE GooglePlaceId = @GooglePlaceId;

            SELECT 'ACTUALIZADO' AS Accion;
            RETURN;
        END

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
            @GooglePlaceId,
            @NombreComplejo,
            @NombreEspacio,
            @CodigoReferencia,
            @CodigoUbigeo,
            @TipoDeporteSuperId,
            @Direccion,
            @Referencia,
            @TelefonoContacto,
            @CorreoContacto,
            @WhatsappContacto,
            COALESCE(@PermiteChatWhatsapp, 0),
            @TarifaReferencial,
            COALESCE(@TieneIluminacion, 0),
            COALESCE(@Techada, 0),
            @GoogleMapsUrl,
            @LatitudReferencia,
            @LongitudReferencia,
            @FotoPrincipalUrl,
            @FotosUrlsCsv,
            COALESCE(@Activo, 1),
            SYSUTCDATETIME(),
            @Usuario
        );

        SELECT 'INSERTADO' AS Accion;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
