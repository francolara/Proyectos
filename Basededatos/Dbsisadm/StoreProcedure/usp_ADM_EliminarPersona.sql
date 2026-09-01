-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/08/2026
-- Description:   Elimina una persona y sus roles de cliente/proveedor solo cuando no tiene operaciones relacionadas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_ADM_EliminarPersona
    @IdEmpresa INT,
    @IdPersona INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Persona AS p
            WHERE p.IdPersona = @IdPersona
              AND p.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR (N'La persona no existe en la empresa activa.', 16, 1);
        END;

        BEGIN TRANSACTION;

        DELETE pr
        FROM dbo.ADM_Proveedor AS pr
        WHERE pr.IdEmpresa = @IdEmpresa
          AND pr.IdPersona = @IdPersona;

        DELETE c
        FROM dbo.ADM_Cliente AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.IdPersona = @IdPersona;

        DELETE p
        FROM dbo.ADM_Persona AS p
        WHERE p.IdPersona = @IdPersona
          AND p.IdEmpresa = @IdEmpresa;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        IF ERROR_NUMBER() = 547
        BEGIN
            SET @ErrorMessage = N'No se puede eliminar la persona porque fue utilizada en compras, ventas, movimientos bancarios, aplicaciones u otra operacion.';
            SET @ErrorSeverity = 16;
            SET @ErrorState = 1;
        END
        ELSE
        BEGIN
            SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        END;

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH;
END;
