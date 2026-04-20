USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/04/2026
-- Description:   Lista niveles activos de desafios para combos.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Desafios_Niveles_Listar]
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            nd.IdNivel,
            nd.Nombre
        FROM dbo.NivelDesafio nd
        WHERE nd.Activo = 1
        ORDER BY nd.Orden, nd.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
