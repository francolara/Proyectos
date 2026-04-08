USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Clientes_Listar para mostrar tipo de documento desde catalogo SUNAT.
-- Firma: Codex - 06/04/2026 | Se agrega filtro opcional por estado activo/inactivo para consulta backend.
-- Firma: Codex - 06/04/2026 | Se elimina dependencia de NegocioClientes y se usa Clientes.NegocioId.
-- Firma: Codex - 07/04/2026 | Se agrega paginacion backend, filtro de busqueda y KPIs globales (activos/inactivos) para listado robusto.
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Listar
    @NegocioId INT,
    @Activo BIT = NULL,
    @Buscar NVARCHAR(200) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT,
    @TotalActivos INT OUTPUT,
    @TotalInactivos INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;
        SET @Buscar = NULLIF(LTRIM(RTRIM(@Buscar)), N'');

        SELECT
            @TotalActivos = COUNT(1)
        FROM dbo.Clientes c
        WHERE c.NegocioId = @NegocioId
          AND c.Activo = 1;

        SELECT
            @TotalInactivos = COUNT(1)
        FROM dbo.Clientes c
        WHERE c.NegocioId = @NegocioId
          AND c.Activo = 0;

        SELECT
            @TotalRegistros = COUNT(1)
        FROM dbo.Clientes c
        WHERE c.NegocioId = @NegocioId
          AND (@Activo IS NULL OR c.Activo = @Activo)
          AND (
                @Buscar IS NULL
                OR c.NombresORazonSocial LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.NombreEquipo, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.NumeroDocumento, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.Telefono, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.Correo, N'') LIKE N'%' + @Buscar + N'%'
              );

        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.NombreEquipo,
            COALESCE(td.CodigoInterno, c.TipoDocumento) AS TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.Activo
        FROM dbo.Clientes c
        LEFT JOIN dbo.TiposDocumentoIdentidadSunat td ON td.CodigoSunat = c.TipoDocumento
        WHERE c.NegocioId = @NegocioId
          AND (@Activo IS NULL OR c.Activo = @Activo)
          AND (
                @Buscar IS NULL
                OR c.NombresORazonSocial LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.NombreEquipo, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.NumeroDocumento, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.Telefono, N'') LIKE N'%' + @Buscar + N'%'
                OR ISNULL(c.Correo, N'') LIKE N'%' + @Buscar + N'%'
              )
        ORDER BY c.NombresORazonSocial, c.Id
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
