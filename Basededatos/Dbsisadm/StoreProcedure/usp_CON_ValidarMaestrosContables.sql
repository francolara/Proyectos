-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Detecta referencias invalidas en los maestros usados por las cargas contables iniciales.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   31/08/2026
-- Description:   Valida las cuentas asignadas a todos los parametros de tipo CONTABLE.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ValidarMaestrosContables
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Incidencias TABLE
        (
            TipoMaestro NVARCHAR(50) NOT NULL,
            CodigoRegistro NVARCHAR(100) NOT NULL,
            Descripcion NVARCHAR(400) NOT NULL
        );

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Plan de cuentas', cuenta.CodigoCuenta, N'La cuenta padre no existe.'
        FROM dbo.CON_PlanCuentaMaestro AS cuenta
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS padre ON padre.CodigoCuenta = cuenta.CodigoCuentaPadre
        WHERE cuenta.CodigoCuentaPadre IS NOT NULL AND padre.IdPlanCuentaMaestro IS NULL;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Cuentas destino', regla.CodigoCuentaOrigen, N'La cuenta origen no existe o no acepta movimiento.'
        FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = regla.CodigoCuentaOrigen AND cuenta.AceptaMovimiento = 1 AND cuenta.Estado = 1
        WHERE cuenta.IdPlanCuentaMaestro IS NULL;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Cuentas destino', regla.CodigoCuentaOrigen, N'El total activo de porcentajes debe ser 100.'
        FROM dbo.CON_CuentaDestinoReglaMaestro AS regla
        LEFT JOIN dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
            ON detalle.IdCuentaDestinoReglaMaestro = regla.IdCuentaDestinoReglaMaestro AND detalle.Activo = 1
        WHERE regla.Activo = 1
        GROUP BY regla.CodigoCuentaOrigen
        HAVING ISNULL(SUM(detalle.Porcentaje), 0) <> 100;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Cuentas destino', CAST(detalle.IdCuentaDestinoReglaDetalleMaestro AS NVARCHAR(100)),
               N'Una cuenta destino de cargo o abono no existe o no acepta movimiento.'
        FROM dbo.CON_CuentaDestinoReglaDetalleMaestro AS detalle
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cargo
            ON cargo.CodigoCuenta = detalle.CodigoCuentaDestinoCargo AND cargo.AceptaMovimiento = 1 AND cargo.Estado = 1
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS abono
            ON abono.CodigoCuenta = detalle.CodigoCuentaDestinoAbono AND abono.AceptaMovimiento = 1 AND abono.Estado = 1
        WHERE cargo.IdPlanCuentaMaestro IS NULL OR abono.IdPlanCuentaMaestro IS NULL;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Parametros', parametro.CodigoParametro, N'La cuenta asignada no existe o no acepta movimiento.'
        FROM dbo.ADM_ParametroMaestro AS parametro
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = parametro.ValorParametro AND cuenta.AceptaMovimiento = 1 AND cuenta.Estado = 1
        WHERE parametro.TipoParametro = 'CONTABLE'
          AND NULLIF(LTRIM(RTRIM(parametro.ValorParametro)), N'') IS NOT NULL
          AND cuenta.IdPlanCuentaMaestro IS NULL;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Impuestos', impuesto.CodigoSunat, N'La cuenta asignada no existe o no acepta movimiento.'
        FROM dbo.CON_TipoImpuesto AS impuesto
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
            ON cuenta.CodigoCuenta = impuesto.CodigoCuenta AND cuenta.AceptaMovimiento = 1 AND cuenta.Estado = 1
        WHERE impuesto.CodigoCuenta IS NOT NULL AND cuenta.IdPlanCuentaMaestro IS NULL;

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Documentos', tipo.CodigoTipoComprobante, N'Una cuenta asignada no existe o no acepta movimiento.'
        FROM dbo.ADM_TipoComprobante AS tipo
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS ventaSoles ON ventaSoles.CodigoCuenta = tipo.CodigoCuentaVentaSoles AND ventaSoles.AceptaMovimiento = 1 AND ventaSoles.Estado = 1
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS ventaDolares ON ventaDolares.CodigoCuenta = tipo.CodigoCuentaVentaDolares AND ventaDolares.AceptaMovimiento = 1 AND ventaDolares.Estado = 1
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS compraSoles ON compraSoles.CodigoCuenta = tipo.CodigoCuentaCompraSoles AND compraSoles.AceptaMovimiento = 1 AND compraSoles.Estado = 1
        LEFT JOIN dbo.CON_PlanCuentaMaestro AS compraDolares ON compraDolares.CodigoCuenta = tipo.CodigoCuentaCompraDolares AND compraDolares.AceptaMovimiento = 1 AND compraDolares.Estado = 1
        WHERE (tipo.CodigoCuentaVentaSoles IS NOT NULL AND ventaSoles.IdPlanCuentaMaestro IS NULL)
           OR (tipo.CodigoCuentaVentaDolares IS NOT NULL AND ventaDolares.IdPlanCuentaMaestro IS NULL)
           OR (tipo.CodigoCuentaCompraSoles IS NOT NULL AND compraSoles.IdPlanCuentaMaestro IS NULL)
           OR (tipo.CodigoCuentaCompraDolares IS NOT NULL AND compraDolares.IdPlanCuentaMaestro IS NULL);

        INSERT INTO @Incidencias (TipoMaestro, CodigoRegistro, Descripcion)
        SELECT N'Configuracion', configuracion.ModuloOperacion + N'/' + configuracion.EscenarioOperacion,
               N'El origen asignado no existe o esta inactivo.'
        FROM dbo.CON_ConfiguracionContabilizacionMaestro AS configuracion
        LEFT JOIN dbo.CON_OrigenMaestro AS origen
            ON origen.CodigoOrigen = configuracion.CodigoOrigen AND origen.Estado = 1
        WHERE origen.IdOrigenMaestro IS NULL;

        SELECT
            incidencia.TipoMaestro,
            incidencia.CodigoRegistro,
            incidencia.Descripcion
        FROM @Incidencias AS incidencia
        ORDER BY incidencia.TipoMaestro, incidencia.CodigoRegistro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
