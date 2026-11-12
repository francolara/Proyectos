-- Firma: FRANCO LARA | Create date: 10/11/2026 | Lee exclusivamente los abonos de un subtipo en una planificación pendiente y vigente.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Exporta el número principal de documento de la ficha y exige guardar las incorporaciones pendientes. Exporta solo los trabajadores incluidos.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TelecreditoHYT_ConsultarLote
    @UsuarioId NVARCHAR(128), @PlanId BIGINT, @VersionFila BINARY(8), @Tipo VARCHAR(10)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PAGOS_PERSONAL.CONSULTAR')=0
            OR dbo.fn_HYTPermiso(@UsuarioId,N'PAGOS_PERSONAL.EDITAR')=0
            OR dbo.fn_HYTPermiso(@UsuarioId,N'TRABAJADORES.CUENTA')=0
            THROW 51000,N'No tiene permiso para exportar archivos de abono.',1;
        IF @Tipo IS NULL OR @Tipo NOT IN ('PLANILLA','EXTERNO')
            THROW 51000,N'Selecciona Planilla u Honorarios.',1;
        IF EXISTS (
            SELECT 1 FROM dbo.PlanPagoHYT AS periodo
            INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.Activo=1 AND trabajador.FrecuenciaPago=periodo.Frecuencia
                AND trabajador.FechaIngreso<=periodo.Hasta
            WHERE periodo.PlanId=@PlanId AND periodo.VersionFila=@VersionFila AND periodo.FechaPago IS NULL
                AND NOT EXISTS (SELECT 1 FROM dbo.PlanPagoDetalleHYT AS detalle
                    WHERE detalle.PlanId=periodo.PlanId AND detalle.TrabajadorId=trabajador.TrabajadorId))
            THROW 51000,N'Guarde la planificación para incorporar a los trabajadores pendientes antes de exportar.',1;

        DECLARE @Datos NVARCHAR(MAX)=(
            SELECT periodo.PlanId,
                JSON_QUERY((SELECT cuenta.TipoCuenta,cuenta.NumeroCuenta,cuenta.Moneda
                    FROM dbo.TelecreditoHYT AS cuenta WHERE cuenta.Id=1 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER)) AS Cargo,
                JSON_QUERY((SELECT detalle.TrabajadorId,trabajador.NombreCompleto AS Nombre,trabajador.TipoDocumentoBcp AS TipoDocumento,
                    trabajador.Documento,trabajador.Banco,
                    trabajador.NumeroCuenta AS Cuenta,trabajador.CuentaInterbancaria AS Cci,
                    trabajador.TipoCuentaBcp AS TipoCuenta,trabajador.MonedaCuentaBcp AS Moneda,
                    detalle.Base+detalle.ImporteExtra-detalle.Descuento AS Neto,
                    CONVERT(BIT,CASE WHEN EXISTS (
                        SELECT 1 FROM dbo.MovimientoPersonalHYT AS movimiento
                        WHERE movimiento.TrabajadorId=detalle.TrabajadorId AND movimiento.MonedaCodigo=N'PEN'
                            AND movimiento.Concepto IN ('REMUNERACION','SERVICIO','DESCUENTO','PAGO')
                            AND movimiento.PeriodoDesde<=periodo.Hasta AND movimiento.PeriodoHasta>=periodo.Desde
                    ) THEN 1 ELSE 0 END) AS TieneMovimientos
                    FROM dbo.PlanPagoDetalleHYT AS detalle
                    JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId=detalle.TrabajadorId
                    WHERE detalle.PlanId=periodo.PlanId AND detalle.Tipo=@Tipo AND detalle.Incluido=1
                    ORDER BY detalle.TrabajadorId FOR JSON PATH,INCLUDE_NULL_VALUES)) AS Abonos
            FROM dbo.PlanPagoHYT AS periodo WITH (HOLDLOCK)
            WHERE periodo.PlanId=@PlanId AND periodo.VersionFila=@VersionFila AND periodo.FechaPago IS NULL
            FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        IF @Datos IS NULL OR @Datos=N''
            THROW 51000,N'La planificación cambió o ya está pagada. Vuelve a abrir el período.',1;
        SELECT @Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
