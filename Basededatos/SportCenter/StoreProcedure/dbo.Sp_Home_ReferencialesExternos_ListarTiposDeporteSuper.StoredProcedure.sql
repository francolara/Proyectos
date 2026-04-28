USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Lista tipos de deporte activos del supermaestro para el barrido de referenciales externos.
-- Firma: Codex - 27/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_ReferencialesExternos_ListarTiposDeporteSuper]
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            tsm.Id,
            tsm.Nombre
        FROM dbo.TiposDeporteSuperMaestro tsm
        WHERE tsm.Activo = 1
        ORDER BY tsm.Nombre;
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

