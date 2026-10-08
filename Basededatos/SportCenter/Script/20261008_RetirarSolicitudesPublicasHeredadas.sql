USE [DbSportCenter];
GO

SET NOCOUNT ON;
SET XACT_ABORT ON;
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   08/10/2026
-- Description:   Retira el flujo heredado de solicitudes publicas reemplazado por la creacion directa de reservas.
-- =============================================
BEGIN TRY
    DECLARE @TieneHistoricos BIT = 0;

    IF OBJECT_ID(N'dbo.SolicitudesReservaPublica', N'U') IS NOT NULL
        EXEC sys.sp_executesql
            N'SELECT @Tiene = CASE WHEN EXISTS (SELECT 1 FROM dbo.SolicitudesReservaPublica) THEN 1 ELSE 0 END;',
            N'@Tiene BIT OUTPUT',
            @Tiene = @TieneHistoricos OUTPUT;

    IF @TieneHistoricos = 1
        RAISERROR('Retiro bloqueado: dbo.SolicitudesReservaPublica contiene historicos. Exporte o depure sus filas antes de continuar.', 16, 1);

    BEGIN TRANSACTION;

    DROP PROCEDURE IF EXISTS dbo.Sp_SolicitudesPublicas_Listar;
    DROP PROCEDURE IF EXISTS dbo.Sp_SolicitudesPublicas_ActualizarEstado;
    DROP PROCEDURE IF EXISTS dbo.Sp_SolicitudesPublicas_ConvertirAReserva;
    DROP PROCEDURE IF EXISTS dbo.Sp_Home_ConsultarSolicitudPublica;
    DROP PROCEDURE IF EXISTS dbo.Sp_Home_ObtenerSolicitudParaNotificacion;
    DROP PROCEDURE IF EXISTS dbo.Sp_Home_MarcarSolicitudNotificada;

    DROP TABLE IF EXISTS dbo.SolicitudesReservaPublica;

    DECLARE @ModuloSolicitudesId INT;

    SELECT @ModuloSolicitudesId = m.Id
    FROM dbo.ModulosSistema AS m
    WHERE m.Codigo = N'SOLICITUDES';

    IF @ModuloSolicitudesId IS NOT NULL
    BEGIN
        DELETE FROM dbo.UsuariosNegocioPermiso
        WHERE ModuloSistemaId = @ModuloSolicitudesId;

        DELETE FROM dbo.RolesNegocioPermiso
        WHERE ModuloSistemaId = @ModuloSolicitudesId;

        DELETE FROM dbo.ModulosSistema
        WHERE Id = @ModuloSolicitudesId;
    END;

    COMMIT TRANSACTION;

    SELECT
        N'FLUJO_SOLICITUDES_PUBLICAS_HEREDADO_RETIRADO' AS Resultado,
        SYSUTCDATETIME() AS FechaUtc;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
GO
