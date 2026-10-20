-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista sucursales para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantSucursales_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT S.IdSucursal, P.GlsPersona
        FROM dbo.Sucursales AS S
        INNER JOIN dbo.Personas AS P ON S.IdSucursal = P.IdPersona
        WHERE S.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR P.GlsPersona LIKE '%' + @Busqueda + '%')
        ORDER BY S.IdSucursal;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
