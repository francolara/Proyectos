-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Consulta catálogos internos de moneda y tratamiento IGV.
-- =============================================
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Catálogos de ubigeo para combos dependientes.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantCatalogos
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51200, N'[HYT_PERMISO] El usuario no está habilitado.', 1;
        SELECT moneda.Codigo, CONCAT(moneda.Nombre, N' (', moneda.Codigo, N') ', moneda.Simbolo) AS Nombre
        FROM dbo.Moneda AS moneda ORDER BY CASE moneda.Codigo WHEN N'PEN' THEN 0 ELSE 1 END, moneda.Codigo;
        SELECT tipo.Codigo, tipo.Nombre FROM dbo.TipoIGV AS tipo ORDER BY tipo.Codigo;
        SELECT d.CodigoDepartamento,d.Nombre FROM dbo.UbigeoDepartamentos AS d WHERE d.Activo=1 ORDER BY d.Nombre;
        SELECT p.CodigoProvincia,p.Nombre FROM dbo.UbigeoProvincias AS p WHERE p.Activo=1 ORDER BY p.Nombre;
        SELECT d.CodigoUbigeo,d.Nombre FROM dbo.UbigeoDistritos AS d WHERE d.Activo=1 ORDER BY d.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
