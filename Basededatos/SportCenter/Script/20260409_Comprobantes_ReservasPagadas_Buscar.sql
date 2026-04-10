/*
Firma: Codex - 09/04/2026
Descripcion: Crea/actualiza SP para buscar reservas pagadas sin comprobante activo (evita doble emision).
*/
USE [DbSportCenter]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ReservasPagadas_Buscar
    @NegocioId INT,
    @Buscar NVARCHAR(150) = NULL,
    @ReservaId INT = NULL,
    @Top INT = 40
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @BuscarNorm NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');
        SET @Top = CASE WHEN ISNULL(@Top, 0) < 1 THEN 40 WHEN @Top > 100 THEN 100 ELSE @Top END;

        ;WITH Fuente AS
        (
            SELECT
                r.Id,
                CONCAT(
                    N'#', r.Id,
                    N' - ',
                    c.NombresORazonSocial,
                    CASE
                        WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN N''
                        ELSE CONCAT(N' [', LTRIM(RTRIM(c.NombreEquipo)), N']')
                    END,
                    N' | ',
                    s.Nombre,
                    N' / ',
                    e.Nombre,
                    N' | ',
                    CONVERT(NVARCHAR(10), r.Fecha, 103),
                    N' ',
                    CONVERT(NVARCHAR(5), r.HoraInicio),
                    N'-',
                    CONVERT(NVARCHAR(5), r.HoraFin),
                    N' | Total: ',
                    CONVERT(NVARCHAR(32), CAST(r.Total AS DECIMAL(10,2)))
                ) AS ReservaTexto,
                r.Fecha,
                r.HoraInicio
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
            WHERE s.NegocioId = @NegocioId
              AND r.Estado = 4
              AND NOT EXISTS
              (
                  SELECT 1
                  FROM dbo.ComprobantesElectronicos ce
                  WHERE ce.NegocioId = @NegocioId
                    AND ce.ReservaId = r.Id
                    AND ce.Estado <> 5
              )
              AND
              (
                  (@BuscarNorm IS NOT NULL AND
                   (
                       CONVERT(NVARCHAR(20), r.Id) LIKE N'%' + @BuscarNorm + N'%'
                       OR c.NombresORazonSocial LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.NombreEquipo, N'') LIKE N'%' + @BuscarNorm + N'%'
                       OR s.Nombre LIKE N'%' + @BuscarNorm + N'%'
                       OR e.Nombre LIKE N'%' + @BuscarNorm + N'%'
                       OR CONVERT(NVARCHAR(10), r.Fecha, 103) LIKE N'%' + @BuscarNorm + N'%'
                   ))
                  OR (@ReservaId IS NOT NULL AND r.Id = @ReservaId)
              )
        )
        SELECT TOP (@Top)
            f.Id,
            f.ReservaTexto
        FROM Fuente f
        ORDER BY
            CASE WHEN @ReservaId IS NOT NULL AND f.Id = @ReservaId THEN 0 ELSE 1 END,
            f.Fecha DESC,
            f.HoraInicio DESC,
            f.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
