-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene los datos personales y de ubicacion del proveedor.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedores_Persona_Obtener
    @IdPersona CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            PE.IdPersona,
            PE.GlsPersona,
            PE.TipoPersona,
            PE.Ruc,
            PE.Direccion,
            PE.IdDistrito,
            U.IdDpto,
            U.IdProv,
            PE.Telefonos,
            PE.Mail,
            PE.IdPais
        FROM dbo.Personas AS PE
        INNER JOIN dbo.Ubigeo AS U
            ON PE.IdDistrito = U.IdDistrito
           AND PE.IdPais = U.IdPais
        WHERE PE.IdPersona = @IdPersona;
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
