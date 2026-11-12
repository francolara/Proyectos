-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta trabajadores con saldo de vacaciones y CCI segun permisos.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Agrega filtro por estado antes de paginar y devuelve los conceptos y datos de Telecrédito en la ficha, con permiso para abono.
-- =============================================
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Devuelve nombres y apellidos y permite buscar en ambos órdenes.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Tipo VARCHAR(10) = NULL,
    @Pagina INT = 1,
    @Activo BIT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar trabajadores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA');
        DECLARE @VerVacaciones BIT = dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.CONSULTAR');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT trabajador.TrabajadorId, trabajador.Documento, trabajador.NombreCompleto, trabajador.Nombres, trabajador.Apellidos, trabajador.Tipo,
               trabajador.FechaIngreso, trabajador.Banco, trabajador.CargoId, cargo.Nombre AS CargoNombre,
               CASE WHEN @TrabajadorId IS NOT NULL THEN trabajador.SueldoMensual END AS SueldoMensual,
               CASE WHEN trabajador.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(trabajador.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta = 1
                    THEN trabajador.NumeroCuenta END AS NumeroCuenta,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta = 1
                    THEN trabajador.CuentaInterbancaria END AS CuentaInterbancaria,
               CASE WHEN @TrabajadorId IS NOT NULL THEN trabajador.TipoDocumentoBcp END AS TipoDocumentoBcp,
               CASE WHEN @TrabajadorId IS NOT NULL THEN trabajador.DocumentoBcp END AS DocumentoBcp,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta=1 THEN trabajador.TipoCuentaBcp END AS TipoCuentaBcp,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta=1 THEN trabajador.MonedaCuentaBcp END AS MonedaCuentaBcp,
               trabajador.HoraEntradaHabitual, trabajador.HoraSalidaHabitual, trabajador.RegimenPrevisional,
               trabajador.AdministradoraAFP, trabajador.FrecuenciaPago, trabajador.Activo, trabajador.VersionFila,
               CASE WHEN @VerVacaciones = 1 AND trabajador.Tipo = 'PLANILLA'
                    THEN COALESCE(vacaciones.Pendientes, 0) END AS VacacionesPendientes,
               CASE WHEN @TrabajadorId IS NOT NULL THEN (
                   SELECT asignado.ConceptoId FROM dbo.TrabajadorConceptoHYT AS asignado
                   WHERE asignado.TrabajadorId=trabajador.TrabajadorId FOR JSON PATH) END AS Conceptos,
               COUNT_BIG(1) OVER () AS Total
        FROM dbo.TrabajadorHYT AS trabajador
        LEFT JOIN dbo.CargoTrabajadorHYT AS cargo ON cargo.CargoId = trabajador.CargoId
        OUTER APPLY (
            SELECT SUM(CONVERT(BIGINT, periodo.DiasReconocidos) - COALESCE(salidas.Dias, 0)) AS Pendientes
            FROM dbo.PeriodoVacacionesHYT AS periodo
            OUTER APPLY (
                SELECT SUM(CONVERT(BIGINT, vacacion.Dias)) AS Dias
                FROM dbo.VacacionesHYT AS vacacion
                WHERE vacacion.PeriodoVacacionesId = periodo.PeriodoVacacionesId AND vacacion.Estado = 'APROBADA'
            ) AS salidas
            WHERE @VerVacaciones = 1 AND trabajador.Tipo = 'PLANILLA' AND periodo.TrabajadorId = trabajador.TrabajadorId
        ) AS vacaciones
        WHERE (@TrabajadorId IS NULL OR trabajador.TrabajadorId = @TrabajadorId)
          AND (@Tipo IS NULL OR trabajador.Tipo = @Tipo)
          AND (@Activo IS NULL OR trabajador.Activo = @Activo)
          AND (@SoloActivos = 0 OR trabajador.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, trabajador.NombreCompleto) > 0
               OR (trabajador.Nombres IS NOT NULL AND CHARINDEX(@Buscar, CONCAT(trabajador.Nombres,N' ',trabajador.Apellidos))>0)
               OR CHARINDEX(@Buscar, trabajador.Documento) > 0)
        ORDER BY trabajador.NombreCompleto, trabajador.TrabajadorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
