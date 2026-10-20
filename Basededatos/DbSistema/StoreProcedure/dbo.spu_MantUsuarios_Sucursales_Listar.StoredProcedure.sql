-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista las empresas y sucursales asignadas a un usuario.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUsuarios_Sucursales_Listar
    @IdEmpresa CHAR(2),
    @IdUsuario CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            SE.Item,
            SE.IdEmpresa,
            E.GlsEmpresa,
            SE.IdSucursal,
            P.GlsPersona AS GlsSucursal
        FROM dbo.SucursalesEmpresa AS SE
        INNER JOIN dbo.Empresas AS E
            ON E.IdEmpresa = SE.IdEmpresa
        INNER JOIN dbo.Sucursales AS S
            ON S.IdEmpresa = SE.IdEmpresa
           AND S.IdSucursal = SE.IdSucursal
        INNER JOIN dbo.Personas AS P
            ON P.IdPersona = S.IdSucursal
        WHERE SE.IdEmpresa = @IdEmpresa
          AND SE.IdUsuario = @IdUsuario
        ORDER BY SE.Item;
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
