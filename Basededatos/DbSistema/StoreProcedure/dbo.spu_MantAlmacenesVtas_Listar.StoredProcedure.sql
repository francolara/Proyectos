-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista las asignaciones de almacenes para ventas de una empresa.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantAlmacenesVtas_Listar
    @IdEmpresa CHAR(2)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            AV.Item,
            AV.IdSucursal,
            S.GlsPersona AS GlsSucursal,
            AV.IdAlmacen,
            A.GlsAlmacen
        FROM dbo.AlmacenesVtas AS AV
        INNER JOIN dbo.Personas AS S
            ON AV.IdSucursal = S.IdPersona
        INNER JOIN dbo.Almacenes AS A
            ON AV.IdAlmacen = A.IdAlmacen
           AND AV.IdEmpresa = A.IdEmpresa
        WHERE AV.IdEmpresa = @IdEmpresa
        ORDER BY AV.Item;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
