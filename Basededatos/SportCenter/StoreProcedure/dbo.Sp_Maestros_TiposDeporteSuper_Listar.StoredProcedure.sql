USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Lista tipos de deporte activos del supermaestro (solo nombre).
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Maestros_TiposDeporteSuper_Listar]
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            tdm.Id,
            tdm.Nombre
        FROM dbo.TiposDeporteSuperMaestro tdm
        WHERE tdm.Activo = 1
        ORDER BY tdm.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
