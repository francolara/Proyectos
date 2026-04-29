USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/04/2026
-- Description:   Actualiza referenciales externos para gestion superadmin (nombre, telefono, tipo de deporte, direccion y ubigeo).
-- Firma: Codex - 29/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialesExternos_ActualizarAdmin]
    @Id INT,
    @NombreComplejo NVARCHAR(180),
    @TelefonoContacto NVARCHAR(40) = NULL,
    @TipoDeporteSuperId INT,
    @Direccion NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6),
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @NombreComplejo = NULLIF(LTRIM(RTRIM(@NombreComplejo)), '');
        SET @TelefonoContacto = NULLIF(LTRIM(RTRIM(@TelefonoContacto)), '');
        SET @Direccion = NULLIF(LTRIM(RTRIM(@Direccion)), '');
        SET @CodigoUbigeo = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @Usuario = NULLIF(LTRIM(RTRIM(@Usuario)), '');

        IF @Id IS NULL OR @Id <= 0
            RAISERROR('Id invalido.', 16, 1);

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

        UPDATE dbo.HomeEspaciosReferencialesExternos
        SET
            NombreComplejo = @NombreComplejo,
            TelefonoContacto = @TelefonoContacto,
            TipoDeporteSuperId = @TipoDeporteSuperId,
            Direccion = @Direccion,
            CodigoUbigeo = @CodigoUbigeo,
            UsuarioActualizacion = COALESCE(@Usuario, UsuarioActualizacion),
            FechaActualizacion = SYSUTCDATETIME()
        WHERE Id = @Id;

        IF @@ROWCOUNT = 0
            RAISERROR('Referencial externo no encontrado.', 16, 1);
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
