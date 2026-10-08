
GO
/****** Object:  StoredProcedure [dbo].[Sp_Maestros_TiposSuelo_Actualizar]    Script Date: 6/04/2026 07:00:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 36_Maestros_PorNegocio_MonedasSuper.sql (linea 437)
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Actualiza estado y protege asociaciones utilizadas por espacios del negocio.
-- =============================================
-- Firma: FRANCO LARA - 06/10/2026 | Impide inactivar una asociacion de suelo utilizada por espacios del negocio.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Maestros_TiposSuelo_Actualizar]
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
            FROM dbo.TiposSuelo ts
            INNER JOIN dbo.EspaciosDeportivos e ON e.TipoSueloSuperId = ts.TipoSueloSuperId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId AND s.NegocioId = ts.NegocioId
            WHERE ts.Id = @Id
              AND ts.NegocioId = @NegocioId
        )
            RAISERROR('No se puede inactivar el tipo de suelo porque existen espacios que lo utilizan.', 16, 1);

        UPDATE dbo.TiposSuelo
        SET
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el tipo de suelo para actualizar.', 16, 1);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'MAESTROS',
            @Accion = N'EDIT',
            @Entidad = N'TipoSuelo',
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
