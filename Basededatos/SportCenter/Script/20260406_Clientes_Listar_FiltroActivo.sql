USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 06/04/2026 | Se agrega filtro opcional de estado activo/inactivo en listado de clientes (backend) usando Clientes.NegocioId.
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Listar
    @NegocioId INT,
    @Activo BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
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
        ORDER BY c.NombresORazonSocial;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
