-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista proveedores para el mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantProveedores_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            PR.IdProveedor,
            PE.GlsPersona,
            CASE WHEN PE.TipoPersona = '01001' THEN 'Natural' ELSE 'Juridica' END AS TipoPersona,
            PE.Ruc,
            PE.Direccion + ', ' + U.GlsUbigeo AS Direccion,
            PR.GlsObservacion
        FROM dbo.Proveedores AS PR
        INNER JOIN dbo.Personas AS PE
            ON PR.IdProveedor = PE.IdPersona
        LEFT JOIN dbo.Ubigeo AS U
            ON PE.IdDistrito = U.IdDistrito
           AND PE.IdPais = U.IdPais
        WHERE PR.IdEmpresa = @IdEmpresa
          AND (
                @Busqueda = ''
                OR PR.IdProveedor LIKE '%' + @Busqueda + '%'
                OR PE.GlsPersona LIKE '%' + @Busqueda + '%'
                OR PR.GlsObservacion LIKE '%' + @Busqueda + '%'
                OR PE.Ruc LIKE '%' + @Busqueda + '%'
              )
        ORDER BY PR.IdProveedor;
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
