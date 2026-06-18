-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Lista personas por empresa con filtros operativos y paginacion server-side.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarPersonasPorEmpresa
    @IdEmpresa INT,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @TipoPersona CHAR(1) = NULL,
    @SoloClientes BIT = 0,
    @SoloProveedores BIT = 0,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), '')
        DECLARE @TipoPersonaTrabajo CHAR(1) = NULLIF(UPPER(LTRIM(RTRIM(@TipoPersona))), '')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                p.IdPersona,
                p.IdEmpresa,
                p.TipoPersona,
                p.TipoDocumento,
                td.Nombre AS NombreTipoDocumento,
                p.NumeroDocumento,
                p.NombreCompleto,
                p.CorreoElectronico,
                p.Telefono,
                p.Direccion,
                p.CodigoUbigeo,
                d.Nombre AS Departamento,
                pr.Nombre AS Provincia,
                di.Nombre AS Distrito,
                CAST(CASE WHEN c.IdCliente IS NULL THEN 0 ELSE 1 END AS BIT) AS EsCliente,
                CAST(CASE WHEN pv.IdProveedor IS NULL THEN 0 ELSE 1 END AS BIT) AS EsProveedor,
                p.Estado
            FROM dbo.ADM_Persona AS p
            INNER JOIN dbo.TiposDocumentoIdentidadSunat AS td
                ON td.CodigoSunat = p.TipoDocumento
            LEFT JOIN dbo.UbigeoDistritos AS di
                ON di.CodigoUbigeo = p.CodigoUbigeo
            LEFT JOIN dbo.UbigeoProvincias AS pr
                ON pr.CodigoProvincia = di.CodigoProvincia
            LEFT JOIN dbo.UbigeoDepartamentos AS d
                ON d.CodigoDepartamento = di.CodigoDepartamento
            LEFT JOIN dbo.ADM_Cliente AS c
                ON c.IdPersona = p.IdPersona
               AND c.IdEmpresa = p.IdEmpresa
               AND c.Estado = 1
            LEFT JOIN dbo.ADM_Proveedor AS pv
                ON pv.IdPersona = p.IdPersona
               AND pv.IdEmpresa = p.IdEmpresa
               AND pv.Estado = 1
            WHERE p.IdEmpresa = @IdEmpresa
              AND (@TipoPersonaTrabajo IS NULL OR p.TipoPersona = @TipoPersonaTrabajo)
              AND (@SoloClientes = 0 OR c.IdCliente IS NOT NULL)
              AND (@SoloProveedores = 0 OR pv.IdProveedor IS NOT NULL)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR p.NumeroDocumento LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR p.NombreCompleto LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(p.CorreoElectronico, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR ISNULL(p.Telefono, '') LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdPersona,
            b.IdEmpresa,
            b.TipoPersona,
            b.TipoDocumento,
            b.NombreTipoDocumento,
            b.NumeroDocumento,
            b.NombreCompleto,
            b.CorreoElectronico,
            b.Telefono,
            b.Direccion,
            b.CodigoUbigeo,
            b.Departamento,
            b.Provincia,
            b.Distrito,
            b.EsCliente,
            b.EsProveedor,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY
            b.NombreCompleto ASC,
            b.IdPersona DESC
        OFFSET CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 0 ELSE (@NumeroPaginaTrabajo - 1) * @TamanoPaginaTrabajo END ROWS
        FETCH NEXT CASE WHEN @NumeroPaginaTrabajo IS NULL OR @TamanoPaginaTrabajo IS NULL THEN 2147483647 ELSE @TamanoPaginaTrabajo END ROWS ONLY;

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
