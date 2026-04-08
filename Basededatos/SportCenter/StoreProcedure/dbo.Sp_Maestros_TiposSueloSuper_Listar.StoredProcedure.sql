USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Lista tipos de suelo activos del supermaestro (solo nombre).
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Maestros_TiposSueloSuper_Listar]
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            tsm.Id,
            tsm.Nombre
        FROM dbo.TiposSueloSuperMaestro tsm
        WHERE tsm.Activo = 1
        ORDER BY tsm.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
