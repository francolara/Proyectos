-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista entidades para el mantenimiento de personas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantPersonas_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            P.IdPersona,
            P.GlsPersona,
            CASE WHEN P.TipoPersona = '01001' THEN 'Natural' ELSE 'Juridica' END AS TipoPersona,
            P.Ruc,
            P.Direccion + ', ' + ISNULL(U.GlsUbigeo, '') AS Direccion,
            P.GlsNombreComercial,
            CASE WHEN ISNULL(PR.IdProveedor, '') <> '' THEN '1' ELSE '0' END AS IndProveedor,
            CASE WHEN ISNULL(C.IdCliente, '') <> '' THEN '1' ELSE '0' END AS IndCliente
        FROM dbo.Personas AS P
        LEFT JOIN dbo.Ubigeo AS U
            ON P.IdDistrito = U.IdDistrito
           AND P.IdPais = U.IdPais
        LEFT JOIN dbo.Proveedores AS PR
            ON PR.IdEmpresa = @IdEmpresa
           AND P.IdPersona = PR.IdProveedor
        LEFT JOIN dbo.Clientes AS C
            ON C.IdEmpresa = @IdEmpresa
           AND P.IdPersona = C.IdCliente
        WHERE @Busqueda = ''
           OR P.GlsPersona LIKE '%' + @Busqueda + '%'
           OR P.IdPersona LIKE '%' + @Busqueda + '%'
           OR P.TipoPersona LIKE '%' + @Busqueda + '%'
           OR P.Direccion LIKE '%' + @Busqueda + '%'
           OR P.Ruc LIKE '%' + @Busqueda + '%'
           OR P.GlsNombreComercial LIKE '%' + @Busqueda + '%'
        ORDER BY P.IdPersona;
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
