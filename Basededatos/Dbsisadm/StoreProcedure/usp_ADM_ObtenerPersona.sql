-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Obtiene una persona de la empresa activa con su estado de cliente y proveedor.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ObtenerPersona
    @IdEmpresa INT,
    @IdPersona INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            p.IdPersona,
            p.IdEmpresa,
            p.TipoPersona,
            p.TipoDocumento,
            p.NumeroDocumento,
            p.ApellidoPaterno,
            p.ApellidoMaterno,
            p.Nombres,
            p.RazonSocial,
            p.NombreCompleto,
            p.CorreoElectronico,
            p.Telefono,
            p.Direccion,
            p.CodigoUbigeo,
            di.CodigoDepartamento,
            di.CodigoProvincia,
            CAST(CASE WHEN c.IdCliente IS NULL THEN 0 ELSE 1 END AS BIT) AS EsCliente,
            CAST(CASE WHEN pv.IdProveedor IS NULL THEN 0 ELSE 1 END AS BIT) AS EsProveedor,
            p.Estado
        FROM dbo.ADM_Persona AS p
        LEFT JOIN dbo.UbigeoDistritos AS di
            ON di.CodigoUbigeo = p.CodigoUbigeo
        LEFT JOIN dbo.ADM_Cliente AS c
            ON c.IdPersona = p.IdPersona
           AND c.IdEmpresa = p.IdEmpresa
           AND c.Estado = 1
        LEFT JOIN dbo.ADM_Proveedor AS pv
            ON pv.IdPersona = p.IdPersona
           AND pv.IdEmpresa = p.IdEmpresa
           AND pv.Estado = 1
        WHERE p.IdEmpresa = @IdEmpresa
          AND p.IdPersona = @IdPersona;

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
