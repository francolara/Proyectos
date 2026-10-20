-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene empresa de transporte, persona y ubicación en una llamada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantEmpTrans_Obtener
    @IdEmpresa CHAR(2), @IdEmpTrans CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT E.IdEmpTrans, P.IdPersona, P.GlsPersona, P.Direccion,
               P.IdDistrito, U.IdDpto, U.IdProv
        FROM dbo.EmpTrans AS E
        INNER JOIN dbo.Personas AS P ON P.IdPersona = E.IdEmpTrans
        LEFT JOIN dbo.Ubigeo AS U ON U.IdDistrito = P.IdDistrito
        WHERE E.IdEmpresa = @IdEmpresa AND E.IdEmpTrans = @IdEmpTrans;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
