
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Lista las resenas publicas de un espacio deportivo en orden descendente por fecha de registro.
-- =============================================
-- Firma:         FRANCO LARA - 08/06/2026 | Expone reseñas publicas por espacio para mostrarlas al final del flujo de reserva del home con usuario visible, fecha y comentario.
-- Firma:         FRANCO LARA - 11/06/2026 | Solo expone reseñas activas e incluye la respuesta publica del negocio cuando existe.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Home_EspacioResenasListar]
    @EspacioDeportivoId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            rr.Id,
            rr.ReservaId,
            r.EspacioDeportivoId,
            rr.AliasPublico,
            rr.Comentario,
            rr.FechaCreacion,
            rr.Respuesta
        FROM dbo.ReservasUsuariosPublicosResenas rr
        INNER JOIN dbo.Reservas r ON r.Id = rr.ReservaId
        WHERE r.EspacioDeportivoId = @EspacioDeportivoId
          AND rr.Activo = 1
        ORDER BY rr.FechaCreacion DESC, rr.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
