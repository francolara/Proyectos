-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta obligaciones y polizas con sus trabajadores cubiertos.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObligacionesHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar obligaciones.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        SELECT (
            SELECT @Mes AS Mes,
                   JSON_QUERY((
                       SELECT obligacion.ObligacionId, obligacion.Tipo, obligacion.Entidad, obligacion.Mes,
                              obligacion.Importe, obligacion.Vencimiento, obligacion.FechaPago,
                              obligacion.ReferenciaPago, obligacion.VersionFila,
                              CASE WHEN obligacion.FechaPago IS NOT NULL THEN 'PAGADA'
                                   WHEN obligacion.Vencimiento < @Hoy THEN 'VENCIDA' ELSE 'PENDIENTE' END AS Estado
                       FROM dbo.ObligacionPersonalHYT AS obligacion
                       WHERE obligacion.Mes = @Mes
                       ORDER BY obligacion.Vencimiento, obligacion.Tipo, obligacion.Entidad
                       FOR JSON PATH
                   )) AS Obligaciones,
                   JSON_QUERY((
                       SELECT poliza.PolizaId, poliza.Aseguradora, poliza.NumeroPoliza, poliza.Cobertura,
                              poliza.VigenciaDesde, poliza.VigenciaHasta, poliza.VersionFila,
                              JSON_QUERY(COALESCE((
                                  SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre
                                  FROM dbo.PolizaSctrTrabajadorHYT AS cobertura
                                  INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = cobertura.TrabajadorId
                                  WHERE cobertura.PolizaId = poliza.PolizaId
                                  ORDER BY trabajador.NombreCompleto FOR JSON PATH
                              ), N'[]')) AS Trabajadores,
                              CASE WHEN poliza.VigenciaHasta < @Hoy THEN 'VENCIDA'
                                   WHEN poliza.VigenciaDesde > @Hoy THEN 'PROGRAMADA'
                                   WHEN DATEDIFF(DAY, @Hoy, poliza.VigenciaHasta) <= 30 THEN 'POR_VENCER'
                                   ELSE 'VIGENTE' END AS Estado
                       FROM dbo.PolizaSctrHYT AS poliza
                       ORDER BY poliza.VigenciaHasta DESC, poliza.PolizaId DESC
                       FOR JSON PATH
                   )) AS Polizas,
                   JSON_QUERY(COALESCE((
                       SELECT trabajador.TrabajadorId, trabajador.NombreCompleto AS Nombre
                       FROM dbo.TrabajadorHYT AS trabajador WHERE trabajador.Activo = 1
                       ORDER BY trabajador.NombreCompleto FOR JSON PATH
                   ), N'[]')) AS PersonalActivo
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
