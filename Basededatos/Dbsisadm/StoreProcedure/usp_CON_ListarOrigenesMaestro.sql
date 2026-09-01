-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Lista los origenes contables maestros con filtros y paginacion para SuperAdmin.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarOrigenesMaestro
    @IdOrigenMaestro INT = NULL,
    @TextoBusqueda NVARCHAR(150) = NULL,
    @SoloActivos BIT = 0,
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 20
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Pagina INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE 1 END;
        DECLARE @Tamano INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 20 END;
        DECLARE @Filtro NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');

        ;WITH Base AS
        (
            SELECT
                origen.IdOrigenMaestro,
                origen.CodigoOrigen,
                origen.NombreOrigen,
                origen.ModuloOrigen,
                origen.PermiteRegistroManual,
                origen.Estado,
                origen.Orden
            FROM dbo.CON_OrigenMaestro AS origen
            WHERE (@IdOrigenMaestro IS NULL OR origen.IdOrigenMaestro = @IdOrigenMaestro)
              AND (@SoloActivos = 0 OR origen.Estado = 1)
              AND
              (
                  @Filtro IS NULL
                  OR origen.CodigoOrigen LIKE '%' + @Filtro + '%'
                  OR origen.NombreOrigen LIKE '%' + @Filtro + '%'
                  OR origen.ModuloOrigen LIKE '%' + @Filtro + '%'
              )
        )
        SELECT
            base.IdOrigenMaestro,
            base.CodigoOrigen,
            base.NombreOrigen,
            base.ModuloOrigen,
            base.PermiteRegistroManual,
            base.Estado,
            base.Orden,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS base
        ORDER BY base.Orden, base.CodigoOrigen
        OFFSET (@Pagina - 1) * @Tamano ROWS
        FETCH NEXT @Tamano ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
