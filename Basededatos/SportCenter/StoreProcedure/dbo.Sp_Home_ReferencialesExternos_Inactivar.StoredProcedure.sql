USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Inactiva un referencial externo del Home desde superadmin.
-- Firma: Codex - 27/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialesExternos_Inactivar]
    @Id INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Usuario = COALESCE(NULLIF(LTRIM(RTRIM(@Usuario)), ''), 'owner-platform');

        IF NOT EXISTS (SELECT 1 FROM dbo.HomeEspaciosReferencialesExternos WHERE Id = @Id)
            RAISERROR('Referencial externo no encontrado.', 16, 1);

        UPDATE dbo.HomeEspaciosReferencialesExternos
           SET Activo = 0,
               FechaActualizacion = SYSUTCDATETIME(),
               UsuarioActualizacion = @Usuario
         WHERE Id = @Id
           AND Activo = 1;
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

