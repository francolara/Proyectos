-- =============================================
-- Author: FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista vehículos para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantVehiculos_Listar @IdEmpresa CHAR(2), @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT V.IdVehiculo, V.GlsVehiculo
        FROM dbo.Vehiculos AS V
        WHERE V.IdEmpresa = @IdEmpresa AND (@Busqueda = '' OR V.GlsVehiculo LIKE '%' + @Busqueda + '%')
        ORDER BY V.IdVehiculo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
