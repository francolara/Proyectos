
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Clientes_ObtenerPorId por integracion de ubigeo fiscal.
-- Firma: Codex - 06/04/2026 | Incluye columnas Nombres y Apellidos para mantenimiento diferenciado de cliente natural/juridico.
-- Firma: Codex - 06/04/2026 | Se elimina dependencia de NegocioClientes y se usa Clientes.NegocioId.
-- Firma: FRANCO LARA - 17/09/2026 | Incluye trazabilidad del cliente para su visualizacion durante la edicion.
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            c.Id,
            c.NombresORazonSocial,
            c.Nombres,
            c.Apellidos,
            c.NombreEquipo,
            c.TipoDocumento,
            c.NumeroDocumento,
            c.Telefono,
            c.Correo,
            c.DireccionFiscal,
            c.CodigoUbigeo,
            c.Activo,
            c.UsuarioCreacion,
            CAST(c.FechaCreacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaRegistro,
            c.UsuarioActualizacion,
            CAST(c.FechaActualizacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaActualizacion
        FROM dbo.Clientes c
        WHERE c.NegocioId = @NegocioId
          AND c.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
