USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Asocia tipo de documento de comprobante al negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_TiposDocumentoComprobante_Crear
    @NegocioId INT,
    @CodigoSunat NVARCHAR(4),
    @Activo BIT = 1,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @CodigoSunat = UPPER(LTRIM(RTRIM(@CodigoSunat)));
        IF @CodigoSunat IS NULL OR @CodigoSunat = N''
            RAISERROR('Selecciona un tipo de documento valido.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoComprobanteSuperMaestro WHERE CodigoSunat = @CodigoSunat AND Activo = 1)
            RAISERROR('El tipo de documento no existe en el supermaestro.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.NegociosTiposDocumentoComprobante WHERE NegocioId = @NegocioId AND CodigoSunat = @CodigoSunat)
            RAISERROR('El tipo de documento ya esta registrado para el negocio.', 16, 1);

        INSERT INTO dbo.NegociosTiposDocumentoComprobante
        (
            NegocioId, CodigoSunat, Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @CodigoSunat, @Activo, SYSUTCDATETIME(), @Usuario
        );

        SELECT SCOPE_IDENTITY();
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
