
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Actualiza estado y protege asociaciones utilizadas por espacios del negocio.
-- =============================================
-- Firma: FRANCO LARA - 06/10/2026 | Impide inactivar una asociacion de deporte utilizada por espacios del negocio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Maestros_TiposDeporte_Actualizar]
    @NegocioId INT,
    @Id INT,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Activo = 0 AND EXISTS
        (
            SELECT 1
            FROM dbo.TiposDeporte td
            INNER JOIN dbo.EspaciosDeportivos e ON e.TipoDeporteSuperId = td.TipoDeporteSuperId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId AND s.NegocioId = td.NegocioId
            WHERE td.Id = @Id
              AND td.NegocioId = @NegocioId
        )
            RAISERROR('No se puede inactivar el deporte porque existen espacios que lo utilizan.', 16, 1);

        UPDATE dbo.TiposDeporte
        SET
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de deporte para actualizar.', 16, 1);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'MAESTROS',
            @Accion = N'EDIT',
            @Entidad = N'TipoDeporte',
            @EntidadId = @Id,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
