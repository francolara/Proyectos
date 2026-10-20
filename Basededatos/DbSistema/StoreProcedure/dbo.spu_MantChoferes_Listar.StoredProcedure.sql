-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista choferes con persona y dirección para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantChoferes_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT C.IdChofer, P.GlsPersona,
               CONCAT(P.Direccion, ', ', U.GlsUbigeo) AS Direccion,
               C.NroBrevete
        FROM dbo.Choferes AS C
        INNER JOIN dbo.Personas AS P ON P.IdPersona = C.IdChofer
        INNER JOIN dbo.Ubigeo AS U ON U.IdDistrito = P.IdDistrito
        WHERE C.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR P.GlsPersona LIKE '%' + @Busqueda + '%')
        ORDER BY C.IdChofer;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
