-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/04/2026
-- Description:   Historial de reservas y obtencion puntual para recordatorio manual desde modulo Reservas.
-- Firma:         Codex - 01/04/2026 | Nuevos SP backend-driven para historial en drawer y recordatorio manual por seleccion.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Historial
    @NegocioId INT,
    @ReservaId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            b.FechaRegistro,
            b.Accion,
            COALESCE(NULLIF(LTRIM(RTRIM(b.UsuarioNombre)), N''), b.UsuarioId, N'sistema') AS UsuarioNombre,
            b.DetalleJson
        FROM dbo.BitacoraAuditoria b
        WHERE b.NegocioId = @NegocioId
          AND b.Modulo = N'RESERVAS'
          AND b.Entidad = N'Reserva'
          AND b.EntidadId = CONVERT(NVARCHAR(80), @ReservaId)
        ORDER BY b.FechaRegistro DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_ObtenerParaRecordatorio
    @NegocioId INT,
    @ReservaId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            s.NegocioId,
            c.NombresORazonSocial AS Cliente,
            c.Correo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            scn.CorreoNotificacion,
            scn.WhatsappContacto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND r.Id = @ReservaId
          AND r.Estado IN (1, 2, 3);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
