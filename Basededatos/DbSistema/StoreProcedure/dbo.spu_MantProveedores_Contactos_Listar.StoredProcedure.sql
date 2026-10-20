-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista los contactos asociados a un proveedor.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedores_Contactos_Listar
    @IdEmpresa CHAR(2),
    @IdProveedor CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PE.IdPersona,
            PE.GlsPersona,
            PE.Telefonos,
            PE.Mail
        FROM dbo.ContactosProveedores AS CP
        INNER JOIN dbo.Personas AS PE
            ON CP.IdContacto = PE.IdPersona
        WHERE CP.IdEmpresa = @IdEmpresa
          AND CP.IdProveedor = @IdProveedor
        ORDER BY PE.GlsPersona;
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
