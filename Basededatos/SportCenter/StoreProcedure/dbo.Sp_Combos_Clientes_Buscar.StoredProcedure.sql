USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 07/04/2026 | Combo de clientes para reservas con busqueda incremental por texto y opcion de incluir cliente especifico.
CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Clientes_Buscar
    @NegocioId INT,
    @Buscar NVARCHAR(150) = NULL,
    @ClienteId INT = NULL,
    @Top INT = 50
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @BuscarNorm NVARCHAR(150) = NULLIF(LTRIM(RTRIM(@Buscar)), N'');
        SET @Top = CASE WHEN ISNULL(@Top, 0) < 1 THEN 50 WHEN @Top > 100 THEN 100 ELSE @Top END;

        ;WITH Fuente AS
        (
            SELECT
                c.Id,
                CONCAT(
                    c.NombresORazonSocial,
                    N' (',
                    c.NumeroDocumento,
                    N')',
                    CASE
                        WHEN NULLIF(LTRIM(RTRIM(c.NombreEquipo)), N'') IS NULL THEN N''
                        ELSE CONCAT(N' - Equipo: ', LTRIM(RTRIM(c.NombreEquipo)))
                    END
                ) AS NombreCliente
            FROM dbo.Clientes c
            WHERE c.NegocioId = @NegocioId
              AND c.Activo = 1
              AND
              (
                  (@BuscarNorm IS NOT NULL AND
                   (
                       c.NombresORazonSocial LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.NumeroDocumento, N'') LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.NombreEquipo, N'') LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.Telefono, N'') LIKE N'%' + @BuscarNorm + N'%'
                       OR ISNULL(c.Correo, N'') LIKE N'%' + @BuscarNorm + N'%'
                   ))
                  OR (@ClienteId IS NOT NULL AND c.Id = @ClienteId)
              )
        )
        SELECT TOP (@Top)
            f.Id,
            f.NombreCliente
        FROM Fuente f
        ORDER BY
            CASE WHEN @ClienteId IS NOT NULL AND f.Id = @ClienteId THEN 0 ELSE 1 END,
            f.NombreCliente ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
