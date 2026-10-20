-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista usuarios por empresa y texto de búsqueda.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUsuarios_Listar
    @IdEmpresa CHAR(2),
    @Texto VARCHAR(100) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            U.IdUsuario,
            P.GlsPersona,
            CONCAT(P.Direccion, ', ', UB.GlsUbigeo) AS Direccion
        FROM dbo.Usuarios AS U
        INNER JOIN dbo.Personas AS P
            ON P.IdPersona = U.IdUsuario
        INNER JOIN dbo.Ubigeo AS UB
            ON UB.IdDistrito = P.IdDistrito
        WHERE U.IdEmpresa = @IdEmpresa
          AND (LTRIM(RTRIM(ISNULL(@Texto, ''))) = '' OR P.GlsPersona LIKE '%' + @Texto + '%')
        ORDER BY U.IdUsuario;
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
