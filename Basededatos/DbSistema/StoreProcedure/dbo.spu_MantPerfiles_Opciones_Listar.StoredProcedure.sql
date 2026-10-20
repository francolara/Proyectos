-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista opciones del sistema e indica las asignadas al perfil.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantPerfiles_Opciones_Listar
    @IdEmpresa CHAR(2),
    @CodSistema CHAR(2),
    @IdPerfil CHAR(8) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            OM.OpmNum,
            OM.OpmDes,
            OM.OpmCod,
            CONVERT(BIT, CASE WHEN EXISTS
            (
                SELECT 1
                FROM dbo.OpcionesPerfil AS OP
                WHERE OP.IdEmpresa = @IdEmpresa
                  AND OP.CodSistema = @CodSistema
                  AND OP.IdPerfil = @IdPerfil
                  AND OP.OpmNum = OM.OpmNum
            ) THEN 1 ELSE 0 END) AS Seleccionado
        FROM dbo.OpcionesMenu AS OM
        WHERE OM.OpmEstado = 'S'
          AND OM.CodSistema = @CodSistema
        ORDER BY OM.OpmNum;
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
