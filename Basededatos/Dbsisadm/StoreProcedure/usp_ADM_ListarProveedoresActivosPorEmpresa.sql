-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista los proveedores activos por empresa con su informacion de persona.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Amplia la salida con correo y telefono para ayudas operativas de compras.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarProveedoresActivosPorEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdProveedor,
            p.IdEmpresa,
            p.IdPersona,
            p.CodigoProveedor,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto,
            pe.CorreoElectronico,
            pe.Telefono,
            p.Contacto,
            p.CuentaDetraccion,
            p.Observacion,
            p.Estado
        FROM dbo.ADM_Proveedor AS p
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = p.IdPersona
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.Estado = 1
          AND pe.Estado = 1
        ORDER BY
            pe.NombreCompleto ASC;

    END TRY

    BEGIN CATCH

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
