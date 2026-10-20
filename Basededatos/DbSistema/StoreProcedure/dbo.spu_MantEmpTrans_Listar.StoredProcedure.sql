-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista empresas de transporte con persona y dirección.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantEmpTrans_Listar
    @IdEmpresa CHAR(2), @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));
        SELECT E.IdEmpTrans, P.GlsPersona,
               CONCAT(P.Direccion, ', ', U.GlsUbigeo) AS Direccion
        FROM dbo.EmpTrans AS E
        INNER JOIN dbo.Personas AS P ON P.IdPersona = E.IdEmpTrans
        INNER JOIN dbo.Ubigeo AS U ON U.IdDistrito = P.IdDistrito
        WHERE E.IdEmpresa = @IdEmpresa
          AND (@Busqueda = '' OR P.GlsPersona LIKE '%' + @Busqueda + '%')
        ORDER BY E.IdEmpTrans;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
