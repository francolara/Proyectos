
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/06/2026
-- Description:   Lista espacios activos de la misma sede para configurarlos como espacios compartidos.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_Combos_EspaciosCompartibles]
    @NegocioId INT,
    @SedeId INT,
    @EspacioActualId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            CONCAT(
                COALESCE(NULLIF(LTRIM(RTRIM(e.Codigo)), N''), N'S/C'),
                N' - ',
                e.Nombre,
                N' (',
                COALESCE(ts.Nombre, N'Sin suelo'),
                N')'
            )
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND e.SedeId = @SedeId
          AND e.Estado = 1
          AND (@EspacioActualId IS NULL OR e.Id <> @EspacioActualId)
        ORDER BY e.Codigo, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
