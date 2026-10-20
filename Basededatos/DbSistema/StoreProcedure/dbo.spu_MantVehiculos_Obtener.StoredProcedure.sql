-- =============================================
-- Author: FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un vehículo para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantVehiculos_Obtener @IdEmpresa CHAR(2), @IdVehiculo CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT V.IdVehiculo,V.GlsVehiculo,V.GlsPlaca,V.IdMarcaVehi,V.GlsModelo,V.GlsColor,V.GlsCodInscripcion,V.IdEmpTrans,V.IdChofer
        FROM dbo.Vehiculos AS V WHERE V.IdEmpresa=@IdEmpresa AND V.IdVehiculo=@IdVehiculo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
