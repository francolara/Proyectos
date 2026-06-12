GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   11/06/2026
-- Description:   Actualiza la respuesta y visibilidad publica de una reseña asociada a un espacio del negocio.
-- =============================================
-- Firma:         FRANCO LARA - 11/06/2026 | Centraliza la moderacion de reseñas para responderlas o retirarlas de la reserva publica sin eliminar el historial.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_ResenaGestionar]
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @ResenaId INT,
    @Respuesta NVARCHAR(800) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @RespuestaNorm NVARCHAR(800) = NULLIF(LTRIM(RTRIM(@Respuesta)), N'');

        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('El negocio es obligatorio.', 16, 1);

        IF @EspacioDeportivoId IS NULL OR @EspacioDeportivoId <= 0
            RAISERROR('El espacio deportivo es obligatorio.', 16, 1);

        IF @ResenaId IS NULL OR @ResenaId <= 0
            RAISERROR('La reseña es obligatoria.', 16, 1);

        IF NULLIF(LTRIM(RTRIM(@Usuario)), N'') IS NULL
            RAISERROR('El usuario de auditoria es obligatorio.', 16, 1);

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ReservasUsuariosPublicosResenas rr
            INNER JOIN dbo.Reservas r ON r.Id = rr.ReservaId
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE rr.Id = @ResenaId
              AND e.Id = @EspacioDeportivoId
              AND s.NegocioId = @NegocioId
        )
            RAISERROR('La reseña no pertenece al espacio deportivo seleccionado.', 16, 1);

        UPDATE rr
        SET
            rr.Respuesta = @RespuestaNorm,
            rr.Activo = @Activo
        FROM dbo.ReservasUsuariosPublicosResenas rr
        WHERE rr.Id = @ResenaId;
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
