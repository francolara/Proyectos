GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/06/2026
-- Description:   Obtiene el detalle de una reserva publica del usuario autenticado para exportarla a calendario ICS.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_UsuariosPublicos_ReservaCalendarioObtener]
    @UsuarioId NVARCHAR(450),
    @ReservaId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF ISNULL(@ReservaId, 0) <= 0
        BEGIN
            RAISERROR(N'La reserva indicada no es valida.', 16, 1);
            RETURN;
        END;

        SELECT TOP (1)
            r.Id AS ReservaId,
            r.Estado AS EstadoId,
            CASE r.Estado
                WHEN 1 THEN N'Reservada'
                WHEN 2 THEN N'Confirmada'
                WHEN 3 THEN N'Pagada'
                WHEN 4 THEN N'Completada'
                WHEN 5 THEN N'Cancelada'
                WHEN 6 THEN N'No Show'
                ELSE N'Pendiente'
            END AS EstadoTexto,
            n.NombreComercial AS NegocioNombre,
            s.Nombre AS SedeNombre,
            e.Nombre AS EspacioNombre,
            s.Direccion AS SedeDireccion,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin
        FROM dbo.ReservasUsuariosPublicos rup
        INNER JOIN dbo.Reservas r ON r.Id = rup.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE rup.UsuarioId = @UsuarioId
          AND rup.ReservaId = @ReservaId;
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
GO
