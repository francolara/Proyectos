-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Lista el historial paginado de exportaciones de libros electrónicos por empresa, periodo y libro.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_Historial_Listar
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @LibroElectronico VARCHAR(10) = NULL,
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 10
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @LibroTrabajo VARCHAR(10) = NULLIF(LTRIM(RTRIM(@LibroElectronico)), '');
        DECLARE @PaginaTrabajo INT = CASE WHEN @NumeroPagina < 1 THEN 1 ELSE @NumeroPagina END;
        DECLARE @TamanoTrabajo INT = CASE WHEN @TamanoPagina < 1 THEN 10 ELSE @TamanoPagina END;

        ;WITH Historial AS
        (
            SELECT
                h.IdLibroElectronicoGeneracion,
                h.IdEmpresa,
                h.Periodo,
                h.CodigoLibro,
                h.CodigoFormato,
                h.NombreArchivo,
                h.CantidadRegistros,
                h.TotalDebe,
                h.TotalHaber,
                h.Estado,
                h.Observaciones,
                h.FechaGeneracion,
                h.UsuarioGeneracion,
                ROW_NUMBER() OVER (ORDER BY h.FechaGeneracion DESC, h.IdLibroElectronicoGeneracion DESC) AS RowNum,
                COUNT(1) OVER () AS TotalRegistros
            FROM dbo.CON_LibroElectronicoGeneracion AS h
            WHERE h.IdEmpresa = @IdEmpresa
              AND h.Periodo = @Periodo
              AND (@LibroTrabajo IS NULL OR h.CodigoLibro = @LibroTrabajo)
        )
        SELECT
            IdLibroElectronicoGeneracion,
            IdEmpresa,
            Periodo,
            CodigoLibro,
            CodigoFormato,
            NombreArchivo,
            CantidadRegistros,
            TotalDebe,
            TotalHaber,
            Estado,
            Observaciones,
            FechaGeneracion,
            UsuarioGeneracion,
            TotalRegistros
        FROM Historial
        WHERE RowNum BETWEEN ((@PaginaTrabajo - 1) * @TamanoTrabajo) + 1 AND @PaginaTrabajo * @TamanoTrabajo
        ORDER BY RowNum;

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
