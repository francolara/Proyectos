-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta ajustes y datos bancarios autorizados del reporte.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Incluye la confirmación de pago y precarga conceptos habituales, divide la base quincenal y aplica asignación familiar y pensiones en la segunda quincena. Devuelve banco y cuenta a usuarios autorizados en el listado y muestra el CCI en abonos interbancarios.
-- =============================================
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Muestra el CCI para abonos fuera de BCP y actualiza el nombre de la ficha en períodos pendientes. Devuelve el cargo actual y precarga trabajadores elegibles que aún no están en una planificación pendiente. Devuelve su inclusión y filtra los excluidos en el reporte.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE,
    @Reporte BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar la planificacion.', 1;

        IF @Frecuencia IS NULL OR @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
           OR @Desde IS NULL OR @Desde > '99991224'
           OR (@Frecuencia = 'SEMANAL' AND DATEDIFF(DAY, '19000101', @Desde) % 7 <> 0)
           OR (@Frecuencia = 'QUINCENAL' AND DAY(@Desde) NOT IN (1,16))
           OR (@Frecuencia = 'MENSUAL' AND DAY(@Desde) <> 1)
            THROW 51000, N'Seleccione el inicio valido del periodo.', 1;

        DECLARE @Hasta DATE = CASE WHEN @Frecuencia = 'SEMANAL' THEN DATEADD(DAY, 6, @Desde)
            WHEN @Frecuencia = 'QUINCENAL' AND DAY(@Desde) = 1 THEN DATEADD(DAY, 14, @Desde)
            ELSE EOMONTH(@Desde) END;

        DECLARE @Quincena TINYINT=CASE WHEN @Frecuencia='QUINCENAL' THEN CASE DAY(@Desde) WHEN 1 THEN 1 ELSE 2 END END;
        DECLARE @PlanId BIGINT, @Version BINARY(8), @FechaPago DATE, @ReferenciaPago NVARCHAR(100);
        SELECT @PlanId = periodo.PlanId, @Version = periodo.VersionFila,
               @FechaPago = periodo.FechaPago, @ReferenciaPago = periodo.ReferenciaPago
        FROM dbo.PlanPagoHYT AS periodo WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;

        DECLARE @VerCuenta BIT = CASE WHEN dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 1 THEN 1 ELSE 0 END;
        DECLARE @Filas TABLE (TrabajadorId BIGINT PRIMARY KEY, Nombre NVARCHAR(180), Cargo NVARCHAR(100),
            Tipo VARCHAR(10), Base DECIMAL(18,2), MinutosExtra INT, ImporteExtra DECIMAL(18,2), Descuento DECIMAL(18,2),
            Ajustes NVARCHAR(MAX), SueldoMensual DECIMAL(18,2), Banco NVARCHAR(80), Cuenta NVARCHAR(34), Cci NVARCHAR(34),
            Incluido BIT NOT NULL DEFAULT (1));
        DECLARE @PorIncorporar BIT = 0;
        IF @PlanId IS NOT NULL
        BEGIN
            INSERT INTO @Filas (TrabajadorId, Nombre, Cargo, Tipo, Base, MinutosExtra, ImporteExtra, Descuento,
                Ajustes, SueldoMensual, Banco, Cuenta, Cci, Incluido)
                SELECT detalle.TrabajadorId,
                       CASE WHEN @FechaPago IS NULL AND trabajador.Nombres IS NOT NULL
                           THEN trabajador.NombreCompleto ELSE detalle.Nombre END AS Nombre,
                       cargo.Nombre AS Cargo, detalle.Tipo, detalle.Base,
                       detalle.MinutosExtra, detalle.ImporteExtra, detalle.Descuento,
                       detalle.Ajustes,
                       CASE WHEN @Quincena=2 THEN COALESCE(trabajador.SueldoMensual,0) END AS SueldoMensual,
                       CASE WHEN @VerCuenta = 1 THEN trabajador.Banco END AS Banco,
                       CASE WHEN @VerCuenta = 1 THEN CASE
                           WHEN @Reporte=0 AND trabajador.CuentaInterbancaria IS NOT NULL
                        AND TRIM(trabajador.Banco) COLLATE Latin1_General_100_CI_AI NOT IN
                            (N'BCP',N'Banco de Crédito',N'Banco de Crédito del Perú',N'Banco de Crédito del Perú (BCP)') THEN trabajador.CuentaInterbancaria
                           ELSE trabajador.NumeroCuenta END END AS Cuenta,
                       CASE WHEN @VerCuenta = 1 AND @Reporte = 1 THEN trabajador.CuentaInterbancaria END AS Cci, detalle.Incluido
                FROM dbo.PlanPagoDetalleHYT AS detalle
                INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = detalle.TrabajadorId
                LEFT JOIN dbo.CargoTrabajadorHYT AS cargo ON cargo.CargoId = trabajador.CargoId
                WHERE detalle.PlanId = @PlanId AND (@Reporte = 0 OR detalle.Incluido = 1);
        END;

        IF @FechaPago IS NULL AND (@PlanId IS NULL OR @Reporte=0)
        BEGIN
            DECLARE @Trabajadores TABLE (TrabajadorId BIGINT PRIMARY KEY, Nombre NVARCHAR(180), Cargo NVARCHAR(100), Tipo VARCHAR(10),
                Regimen VARCHAR(3), AdministradoraAFP NVARCHAR(100), Sueldo DECIMAL(18,2), Base DECIMAL(18,2),
                Banco NVARCHAR(80), Cuenta NVARCHAR(34));
            INSERT INTO @Trabajadores(TrabajadorId,Nombre,Cargo,Tipo,Regimen,AdministradoraAFP,Sueldo,Base,Banco,Cuenta)
            SELECT trabajador.TrabajadorId,trabajador.NombreCompleto,cargo.Nombre,trabajador.Tipo,trabajador.RegimenPrevisional,
                trabajador.AdministradoraAFP,sueldo.Mensual,
                CASE @Quincena WHEN 1 THEN ROUND(sueldo.Mensual/2,2)
                    WHEN 2 THEN sueldo.Mensual-ROUND(sueldo.Mensual/2,2) ELSE sueldo.Mensual END,
                CASE WHEN @VerCuenta=1 THEN trabajador.Banco END,
                CASE WHEN @VerCuenta=1 THEN CASE
                    WHEN @Reporte=0 AND trabajador.CuentaInterbancaria IS NOT NULL
                        AND TRIM(trabajador.Banco) COLLATE Latin1_General_100_CI_AI NOT IN
                            (N'BCP',N'Banco de Crédito',N'Banco de Crédito del Perú',N'Banco de Crédito del Perú (BCP)') THEN trabajador.CuentaInterbancaria
                    ELSE trabajador.NumeroCuenta END END
            FROM dbo.TrabajadorHYT AS trabajador
            LEFT JOIN dbo.CargoTrabajadorHYT AS cargo ON cargo.CargoId = trabajador.CargoId
            CROSS APPLY (VALUES(COALESCE(trabajador.SueldoMensual,0))) AS sueldo(Mensual)
            WHERE trabajador.Activo=1 AND trabajador.FrecuenciaPago=@Frecuencia AND trabajador.FechaIngreso<=@Hasta
                AND NOT EXISTS (SELECT 1 FROM @Filas AS guardado WHERE guardado.TrabajadorId=trabajador.TrabajadorId);
            IF @PlanId IS NOT NULL AND @@ROWCOUNT > 0 SET @PorIncorporar = 1;

            DECLARE @Precarga TABLE (TrabajadorId BIGINT, ConceptoId BIGINT, Nombre NVARCHAR(100),
                Tipo VARCHAR(10), Porcentaje DECIMAL(7,4), Importe DECIMAL(18,2), BasePorcentaje DECIMAL(18,2));
            INSERT INTO @Precarga(TrabajadorId,ConceptoId,Nombre,Tipo,Porcentaje,Importe,BasePorcentaje)
            SELECT personal.TrabajadorId,concepto.ConceptoId,concepto.Nombre,
                CASE WHEN concepto.Naturaleza='ABONO' THEN 'ABONO' ELSE 'DESCUENTO' END,
                CASE WHEN concepto.Calculo='PORCENTAJE' THEN tasa.Valor END,
                CASE WHEN concepto.Calculo='PORCENTAJE' THEN
                    ROUND(COALESCE(calculo.BaseMensual,personal.Base)*COALESCE(tasa.Valor,0)/100,2)
                    ELSE COALESCE(concepto.Valor,0) END,
                calculo.BaseMensual
            FROM @Trabajadores AS personal
            CROSS JOIN dbo.ConceptoPlanillaHYT AS concepto
            LEFT JOIN dbo.EntidadPrevisionalHYT AS entidad ON entidad.Regimen=concepto.Regimen
                AND (entidad.Regimen='ONP' OR entidad.Nombre=personal.AdministradoraAFP)
            CROSS APPLY (VALUES(COALESCE(concepto.Valor,entidad.PorcentajeEstimado))) AS tasa(Valor)
            CROSS APPLY (VALUES(CASE WHEN @Quincena=2 AND concepto.Calculo='PORCENTAJE'
                AND (concepto.Regimen IS NOT NULL OR concepto.CodigoSistema='ASIGNACION_FAMILIAR')
                THEN personal.Sueldo END)) AS calculo(BaseMensual)
            WHERE concepto.Activo=1
                AND (ISNULL(@Quincena,0)<>1 OR (concepto.Regimen IS NULL AND concepto.CodigoSistema IS NULL))
                AND ((concepto.Regimen=personal.Regimen AND personal.Tipo='PLANILLA')
                    OR EXISTS (SELECT 1 FROM dbo.TrabajadorConceptoHYT AS asignado
                        WHERE asignado.TrabajadorId=personal.TrabajadorId AND asignado.ConceptoId=concepto.ConceptoId));

            INSERT INTO @Filas (TrabajadorId, Nombre, Cargo, Tipo, Base, MinutosExtra, Banco, Cuenta,
                SueldoMensual, ImporteExtra, Descuento, Ajustes)
                SELECT personal.TrabajadorId, personal.Nombre, personal.Cargo, personal.Tipo, personal.Base, 0 AS MinutosExtra,
                       personal.Banco,personal.Cuenta,
                       CASE WHEN @Quincena=2 THEN personal.Sueldo END AS SueldoMensual,
                       COALESCE(totales.Abonos,0) AS ImporteExtra, COALESCE(totales.Cargos,0) AS Descuento,
                       JSON_QUERY((SELECT ajuste.Tipo,@Desde AS Fecha,ajuste.Nombre AS Concepto,
                           0 AS Horas,ajuste.Importe,ajuste.ConceptoId,ajuste.Porcentaje,ajuste.BasePorcentaje
                           FROM @Precarga AS ajuste WHERE ajuste.TrabajadorId=personal.TrabajadorId
                           ORDER BY ajuste.Tipo,ajuste.Nombre FOR JSON PATH)) AS Ajustes
                FROM @Trabajadores AS personal
                OUTER APPLY (
                    SELECT SUM(CASE WHEN ajuste.Tipo='ABONO' THEN ajuste.Importe ELSE 0 END) AS Abonos,
                           SUM(CASE WHEN ajuste.Tipo='DESCUENTO' THEN ajuste.Importe ELSE 0 END) AS Cargos
                    FROM @Precarga AS ajuste WHERE ajuste.TrabajadorId=personal.TrabajadorId
                ) AS totales;
        END;

        SELECT (SELECT @PlanId AS PlanId, @Version AS VersionFila, @Frecuencia AS Frecuencia,
            @Desde AS Desde, @Hasta AS Hasta, @VerCuenta AS VerCuenta,
            @FechaPago AS FechaPago, @ReferenciaPago AS ReferenciaPago, @PorIncorporar AS PorIncorporar,
            JSON_QUERY(COALESCE((SELECT fila.TrabajadorId, fila.Nombre, fila.Cargo, fila.Tipo, fila.Base, fila.MinutosExtra,
                fila.ImporteExtra, fila.Descuento, JSON_QUERY(fila.Ajustes) AS Ajustes, fila.SueldoMensual,
                fila.Banco, fila.Cuenta, fila.Cci, fila.Incluido FROM @Filas AS fila
                ORDER BY fila.Tipo, fila.Nombre, fila.TrabajadorId FOR JSON PATH), N'[]')) AS Filas,
            JSON_QUERY(CASE WHEN @Reporte=0 AND @FechaPago IS NULL THEN (
                SELECT concepto.ConceptoId,concepto.Nombre,concepto.Naturaleza,concepto.Calculo,concepto.Valor,
                    CONVERT(BIT,CASE WHEN @Quincena=2 AND (concepto.Regimen IS NOT NULL OR concepto.CodigoSistema='ASIGNACION_FAMILIAR')
                        THEN 1 ELSE 0 END) AS BaseMensual
                FROM dbo.ConceptoPlanillaHYT AS concepto WHERE concepto.Activo=1
                    AND (ISNULL(@Quincena,0)<>1 OR (concepto.Regimen IS NULL AND concepto.CodigoSistema IS NULL))
                ORDER BY concepto.Nombre FOR JSON PATH) END) AS Conceptos
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
