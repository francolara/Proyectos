-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene chofer, persona y ubicación en una sola llamada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantChoferes_Obtener
    @IdEmpresa CHAR(2),
    @IdChofer CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT C.IdChofer, C.NroBrevete, P.IdPersona, P.GlsPersona,
               P.Direccion, P.IdDistrito, U.IdDpto, U.IdProv
        FROM dbo.Choferes AS C
        INNER JOIN dbo.Personas AS P ON P.IdPersona = C.IdChofer
        LEFT JOIN dbo.Ubigeo AS U ON U.IdDistrito = P.IdDistrito
        WHERE C.IdEmpresa = @IdEmpresa AND C.IdChofer = @IdChofer;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
