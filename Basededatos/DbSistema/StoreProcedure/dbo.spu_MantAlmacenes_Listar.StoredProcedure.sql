-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista almacenes para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantAlmacenes_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT A.IdSucursal, S.GlsPersona AS GlsSucursal, A.IdAlmacen, A.GlsAlmacen
        FROM dbo.Almacenes AS A
        INNER JOIN dbo.Personas AS S ON A.IdSucursal = S.IdPersona
        WHERE A.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR A.GlsAlmacen LIKE '%' + @Busqueda + '%')
        ORDER BY A.IdAlmacen;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
