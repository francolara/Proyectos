-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Devuelve bancos y entidades previsionales de referencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CatalogosPlanillasHYT_Consultar
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51000, N'No tiene permiso para consultar catalogos.', 1;
        SELECT (SELECT JSON_QUERY((SELECT banco.Nombre FROM dbo.BancoPeruHYT AS banco ORDER BY banco.Nombre FOR JSON PATH)) AS Bancos,
            JSON_QUERY((SELECT entidad.Nombre, entidad.Regimen FROM dbo.EntidadPrevisionalHYT AS entidad
                ORDER BY entidad.Regimen, entidad.Nombre FOR JSON PATH)) AS Entidades
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
