-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista los contactos asociados a un cliente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Contactos_Listar
    @IdEmpresa CHAR(2),
    @IdCliente CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.IdPersona,
            P.GlsPersona,
            P.Telefonos,
            P.Mail
        FROM dbo.ContactosClientes AS C
        INNER JOIN dbo.Personas AS P
            ON C.IdContacto = P.IdPersona
        WHERE C.IdEmpresa = @IdEmpresa
          AND C.IdCliente = @IdCliente
        ORDER BY P.GlsPersona;
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
