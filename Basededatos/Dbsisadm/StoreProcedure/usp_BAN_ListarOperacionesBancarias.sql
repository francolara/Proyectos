-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Lista operaciones bancarias por destino para el arbol de ingresos y egresos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Permite filtrar operaciones bancarias por tipo operativo para reutilizar el mismo catalogo en transferencias entre cuentas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_ListarOperacionesBancarias
    @TipoMovimiento CHAR(1),
    @TextoBusqueda NVARCHAR(200) = NULL,
    @TamanoPagina INT = 100,
    @IdTipoOpeBancaria CHAR(1) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @TextoBusquedaTrabajo NVARCHAR(200) = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');
        DECLARE @TamanoPaginaTrabajo INT = CASE WHEN ISNULL(@TamanoPagina, 0) > 0 THEN @TamanoPagina ELSE 100 END;
        DECLARE @IdTipoOpeBancariaTrabajo CHAR(1) = NULLIF(LTRIM(RTRIM(@IdTipoOpeBancaria)), '');

        ;WITH Operaciones AS
        (
            SELECT
                LTRIM(RTRIM(op.idOpeBancaria)) AS IdOpeBancaria,
                LTRIM(RTRIM(op.idTipoOpeBancaria)) AS IdTipoOpeBancaria,
                LTRIM(RTRIM(op.Destino)) AS TipoMovimiento,
                MAX(LTRIM(RTRIM(op.Tipo))) AS TipoOperacion
            FROM dbo.operacionesbancarias AS op
            WHERE LTRIM(RTRIM(op.Destino)) = @TipoMovimiento
              AND (@IdTipoOpeBancariaTrabajo IS NULL OR LTRIM(RTRIM(op.idTipoOpeBancaria)) = @IdTipoOpeBancariaTrabajo)
            GROUP BY
                LTRIM(RTRIM(op.idOpeBancaria)),
                LTRIM(RTRIM(op.idTipoOpeBancaria)),
                LTRIM(RTRIM(op.Destino))
        )
        SELECT TOP (@TamanoPaginaTrabajo)
            o.IdOpeBancaria,
            o.IdTipoOpeBancaria,
            o.TipoMovimiento,
            o.TipoOperacion
        FROM Operaciones AS o
        WHERE NULLIF(o.TipoOperacion, '') IS NOT NULL
          AND (
                @TextoBusquedaTrabajo IS NULL
                OR o.IdOpeBancaria LIKE '%' + @TextoBusquedaTrabajo + '%'
                OR o.TipoOperacion LIKE '%' + @TextoBusquedaTrabajo + '%'
              )
        ORDER BY o.TipoOperacion ASC;

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
