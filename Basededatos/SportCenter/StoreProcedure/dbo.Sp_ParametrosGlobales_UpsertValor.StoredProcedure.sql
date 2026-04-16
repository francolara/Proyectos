USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 15/04/2026 | Procedimiento para actualizar solo ValorParametro por NombreParametro en ParametrosGlobales (sin inserts; alta por script).
CREATE OR ALTER PROCEDURE dbo.Sp_ParametrosGlobales_UpsertValor
    @NombreParametro NVARCHAR(100),
    @Descripcion NVARCHAR(250) = NULL,
    @ValorParametro NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @ValorNorm NVARCHAR(100) = LEFT(COALESCE(@ValorParametro, N''), 100);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.ParametrosGlobales p
            WHERE p.NombreParametro = @NombreParametro
        )
        BEGIN
            RAISERROR (N'El parametro global %s no existe. Debe crearse por script.', 16, 1, @NombreParametro);
            RETURN;
        END

        UPDATE p
           SET p.ValorParametro = @ValorNorm
        FROM dbo.ParametrosGlobales p
        WHERE p.NombreParametro = @NombreParametro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
