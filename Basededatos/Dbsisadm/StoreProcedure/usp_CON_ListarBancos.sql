-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Lista bancos del catalogo maestro para ayudas operativas con filtro y paginacion opcional.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarBancos
    @SoloActivos BIT = 1,
    @TextoBusqueda NVARCHAR(200) = NULL,
    @NumeroPagina INT = NULL,
    @TamanoPagina INT = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'')
        DECLARE @NumeroPaginaTrabajo INT = CASE WHEN ISNULL(@NumeroPagina, 0) > 0 THEN @NumeroPagina ELSE NULL END
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE NULL END

        ;WITH Base AS
        (
            SELECT
                b.IdBanco,
                b.Codigo,
                b.Nombre,
                b.Estado
            FROM dbo.CON_Bancos AS b
            WHERE (@SoloActivos = 0 OR b.Estado = 1)
              AND (
                    @TextoBusquedaTrabajo IS NULL
                    OR b.Codigo LIKE '%' + @TextoBusquedaTrabajo + '%'
                    OR b.Nombre LIKE '%' + @TextoBusquedaTrabajo + '%'
                  )
        )
        SELECT
            b.IdBanco,
            b.Codigo,
            b.Nombre,
            b.Estado,
            COUNT(1) OVER() AS TotalRegistros
        FROM Base AS b
        ORDER BY b.Nombre ASC
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
