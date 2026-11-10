-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Calcula el estimado mensual por entidad y su detalle de trabajadores a partir de sueldos de planilla.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_EstimadoPrevisionalHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE,
    @IncluirTrabajadores BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.CONSULTAR') = 0
           OR dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para consultar estimados previsionales.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @FinMes DATE = EOMONTH(@Mes);
        WITH Detalle AS
        (
            SELECT entidad.Nombre AS Entidad, trabajador.Documento, trabajador.NombreCompleto AS Nombre,
                   COALESCE(trabajador.SueldoMensual, 0) AS SueldoMensual,
                   CONVERT(DECIMAL(24,2), COALESCE(trabajador.SueldoMensual, 0) * entidad.PorcentajeEstimado / 100) AS Estimado
            FROM dbo.EntidadPrevisionalHYT AS entidad
            INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.Tipo = 'PLANILLA'
                AND trabajador.Activo = 1 AND trabajador.FechaIngreso <= @FinMes
                AND trabajador.RegimenPrevisional = entidad.Regimen
                AND (entidad.Regimen = 'ONP' OR trabajador.AdministradoraAFP = entidad.Nombre)
            WHERE entidad.PorcentajeEstimado IS NOT NULL
        ),
        Totales AS
        (
            SELECT detalle.Entidad, SUM(detalle.SueldoMensual) AS BaseMensual, SUM(detalle.Estimado) AS Estimado
            FROM Detalle AS detalle
            GROUP BY detalle.Entidad
        )
        SELECT entidad.Nombre AS Entidad, entidad.Regimen, entidad.PorcentajeEstimado AS Porcentaje,
               CONVERT(DECIMAL(24,2), COALESCE(totales.BaseMensual, 0)) AS BaseMensual,
               CONVERT(DECIMAL(24,2), COALESCE(totales.Estimado, 0)) AS Estimado,
               CASE WHEN @IncluirTrabajadores = 1 THEN COALESCE((
                   SELECT detalle.Documento, detalle.Nombre, detalle.SueldoMensual, detalle.Estimado
                   FROM Detalle AS detalle
                   WHERE detalle.Entidad = entidad.Nombre
                   ORDER BY detalle.Nombre, detalle.Documento
                   FOR JSON PATH
               ), N'[]') END AS Trabajadores
        FROM dbo.EntidadPrevisionalHYT AS entidad
        LEFT JOIN Totales AS totales ON totales.Entidad = entidad.Nombre
        WHERE entidad.PorcentajeEstimado IS NOT NULL
        ORDER BY entidad.Regimen, entidad.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
