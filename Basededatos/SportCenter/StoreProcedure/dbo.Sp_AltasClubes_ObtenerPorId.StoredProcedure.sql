USE [dbsportcenter_20260613]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/04/2026
-- Description:   Obtiene una solicitud de alta de club por Id para procesos de notificacion.
-- Firma:         Codex - 26/04/2026 | Nuevo SP para recuperar correo y datos de solicitud al aprobar desde superadmin.
-- Firma:         FRANCO LARA - 21/07/2026 | Expone el plan comercial publico de la solicitud.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[Sp_AltasClubes_ObtenerPorId]
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            ac.Id,
            ac.CodigoSolicitud,
            ac.NombreContacto,
            ac.Telefono,
            ac.Correo,
            ac.RelacionClub,
            ac.NombreClub,
            ac.Pais,
            ac.ProvinciaEstado,
            ac.Ciudad,
            ac.Direccion,
            ac.Estado,
            ac.ComentarioGestion,
            ac.NegocioId,
            ac.SedeId,
            ac.FechaRegistro,
            ac.FechaGestion,
            COALESCE(ac.PlanComercial, N'PRUEBA') AS PlanComercial
        FROM dbo.SolicitudesAltaClub ac
        WHERE ac.Id = @Id;
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
