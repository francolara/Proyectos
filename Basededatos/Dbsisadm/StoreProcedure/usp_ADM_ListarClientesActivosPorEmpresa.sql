-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Lista los clientes activos por empresa con su informacion de persona.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Amplia la salida con correo y telefono para ayudas operativas de ventas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarClientesActivosPorEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdCliente,
            c.IdEmpresa,
            c.IdPersona,
            c.CodigoCliente,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto,
            pe.CorreoElectronico,
            pe.Telefono,
            c.LimiteCredito,
            c.DiasCredito,
            c.Observacion,
            c.Estado
        FROM dbo.ADM_Cliente AS c
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.Estado = 1
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
