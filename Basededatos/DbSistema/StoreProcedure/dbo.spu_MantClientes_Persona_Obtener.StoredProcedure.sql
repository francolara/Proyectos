-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene los datos personales y de ubicacion del cliente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Persona_Obtener
    @IdPersona CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            P.IdPersona,
            P.GlsPersona,
            P.TipoPersona,
            P.Ruc,
            P.Direccion,
            P.IdDistrito,
            U.IdDpto,
            U.IdProv,
            P.Telefonos,
            P.Mail,
            P.GlsContacto,
            P.IdPais
        FROM dbo.Personas AS P
        INNER JOIN dbo.Ubigeo AS U
            ON P.IdDistrito = U.IdDistrito
           AND P.IdPais = U.IdPais
        WHERE P.IdPersona = @IdPersona;
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
