-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda ajustes detallados y calcula los acumulados.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Bloquea la edicion despues de confirmar el pago.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PlanPagoHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Frecuencia VARCHAR(10),
    @Desde DATE,
    @VersionFila BINARY(8),
    @Filas NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar la planificacion.', 1;

        IF @Frecuencia IS NULL OR @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
           OR @Desde IS NULL OR @Desde > '99991224'
           OR (@Frecuencia = 'SEMANAL' AND DATEDIFF(DAY, '19000101', @Desde) % 7 <> 0)
           OR (@Frecuencia = 'QUINCENAL' AND DAY(@Desde) NOT IN (1,16))
           OR (@Frecuencia = 'MENSUAL' AND DAY(@Desde) <> 1)
            THROW 51000, N'Seleccione el inicio valido del periodo.', 1;

        DECLARE @Hasta DATE = CASE WHEN @Frecuencia = 'SEMANAL' THEN DATEADD(DAY, 6, @Desde)
            WHEN @Frecuencia = 'QUINCENAL' AND DAY(@Desde) = 1 THEN DATEADD(DAY, 14, @Desde)
            ELSE EOMONTH(@Desde) END;

        IF COALESCE(ISJSON(@Filas, ARRAY), 0) <> 1
            THROW 51000, N'La planificacion no es valida.', 1;

        DECLARE @Datos TABLE (TrabajadorId BIGINT PRIMARY KEY, Base DECIMAL(20,4), MinutosExtra INT,
                              ImporteExtra DECIMAL(20,4), Descuento DECIMAL(20,4), Ajustes NVARCHAR(MAX));
        INSERT INTO @Datos (TrabajadorId, Base, MinutosExtra, ImporteExtra, Descuento, Ajustes)
        SELECT fila.TrabajadorId, fila.Base, fila.MinutosExtra, fila.ImporteExtra, fila.Descuento, fila.Ajustes
        FROM OPENJSON(@Filas) WITH (TrabajadorId BIGINT, Base DECIMAL(20,4), MinutosExtra INT,
            ImporteExtra DECIMAL(20,4), Descuento DECIMAL(20,4), Ajustes NVARCHAR(MAX) AS JSON) AS fila;
        IF EXISTS (SELECT 1 FROM @Datos AS fila WHERE fila.Ajustes IS NOT NULL AND ISJSON(fila.Ajustes, ARRAY) <> 1)
            THROW 51000, N'Los ajustes del periodo no son validos.', 1;

        DECLARE @Ajustes TABLE (TrabajadorId BIGINT, Tipo VARCHAR(10), Fecha DATE, Presupuesto NVARCHAR(40),
            Concepto NVARCHAR(300), Horas INT, Importe DECIMAL(20,4));
        INSERT INTO @Ajustes (TrabajadorId, Tipo, Fecha, Presupuesto, Concepto, Horas, Importe)
        SELECT fila.TrabajadorId, ajuste.Tipo, ajuste.Fecha, NULLIF(TRIM(ajuste.Presupuesto), N''),
               NULLIF(TRIM(ajuste.Concepto), N''), ajuste.Horas, ajuste.Importe
        FROM @Datos AS fila CROSS APPLY OPENJSON(fila.Ajustes)
        WITH (Tipo VARCHAR(10), Fecha DATE, Presupuesto NVARCHAR(40), Concepto NVARCHAR(300),
              Horas INT, Importe DECIMAL(20,4)) AS ajuste;

        IF EXISTS (
            SELECT 1 FROM @Ajustes AS ajuste
            WHERE ajuste.Tipo IS NULL OR ajuste.Tipo NOT IN ('EXTRA', 'DESCUENTO')
               OR ajuste.Fecha IS NULL OR ajuste.Fecha NOT BETWEEN @Desde AND @Hasta
               OR ajuste.Horas IS NULL OR (ajuste.Tipo = 'EXTRA' AND ajuste.Horas NOT BETWEEN 1 AND 24)
               OR (ajuste.Tipo = 'DESCUENTO' AND (ajuste.Horas <> 0 OR ajuste.Concepto IS NULL))
               OR (ajuste.Presupuesto IS NULL AND ajuste.Concepto IS NULL)
               OR ajuste.Importe IS NULL OR ajuste.Importe < 0 OR ajuste.Importe <> ROUND(ajuste.Importe, 2)
               OR (ajuste.Tipo = 'DESCUENTO' AND ajuste.Importe = 0)
               OR (ajuste.Presupuesto IS NOT NULL AND NOT EXISTS (
                   SELECT 1 FROM dbo.Presupuesto AS presupuesto
                   WHERE presupuesto.NumeroPresupuesto = ajuste.Presupuesto OR presupuesto.NumeroInterno = ajuste.Presupuesto)))
            THROW 51000, N'Revise fecha, presupuesto, concepto, horas e importe de los ajustes.', 1;

        UPDATE fila SET MinutosExtra = COALESCE(totales.Horas, 0) * 60,
            ImporteExtra = COALESCE(totales.Extra, 0), Descuento = COALESCE(totales.Descuento, 0)
        FROM @Datos AS fila
        OUTER APPLY (
            SELECT SUM(CASE WHEN ajuste.Tipo = 'EXTRA' THEN ajuste.Horas ELSE 0 END) AS Horas,
                   SUM(CASE WHEN ajuste.Tipo = 'EXTRA' THEN ajuste.Importe ELSE 0 END) AS Extra,
                   SUM(CASE WHEN ajuste.Tipo = 'DESCUENTO' THEN ajuste.Importe ELSE 0 END) AS Descuento
            FROM @Ajustes AS ajuste WHERE ajuste.TrabajadorId = fila.TrabajadorId
        ) AS totales
        WHERE fila.Ajustes IS NOT NULL;

        IF NOT EXISTS (SELECT 1 FROM @Datos) OR EXISTS (
            SELECT 1 FROM @Datos AS fila
            WHERE fila.Base IS NULL OR fila.ImporteExtra IS NULL OR fila.Descuento IS NULL OR fila.MinutosExtra IS NULL
               OR fila.Base < 0 OR fila.ImporteExtra < 0 OR fila.Descuento < 0 OR fila.MinutosExtra < 0
               OR fila.Descuento > fila.Base + fila.ImporteExtra
               OR fila.Base <> ROUND(fila.Base, 2) OR fila.ImporteExtra <> ROUND(fila.ImporteExtra, 2)
               OR fila.Descuento <> ROUND(fila.Descuento, 2))
            THROW 51000, N'Revise los importes y minutos de la planificacion.', 1;

        BEGIN TRANSACTION;
        DECLARE @PlanId BIGINT, @Actual BINARY(8), @FechaPago DATE;
        SELECT @PlanId = periodo.PlanId, @Actual = periodo.VersionFila, @FechaPago = periodo.FechaPago
        FROM dbo.PlanPagoHYT AS periodo WITH (UPDLOCK, HOLDLOCK)
        WHERE periodo.Frecuencia = @Frecuencia AND periodo.Desde = @Desde;
        IF @FechaPago IS NOT NULL
            THROW 51000, N'La planificacion ya esta pagada y cerrada.', 1;
        IF @Actual IS DISTINCT FROM @VersionFila
            THROW 51000, N'La planificacion cambio. Vuelva a abrirla.', 1;

        IF @PlanId IS NULL
        BEGIN
            IF EXISTS (
                SELECT 1 FROM @Datos AS fila
                LEFT JOIN dbo.TrabajadorHYT AS trabajador WITH (HOLDLOCK) ON trabajador.TrabajadorId = fila.TrabajadorId
                WHERE trabajador.TrabajadorId IS NULL OR trabajador.Activo = 0
                   OR trabajador.FrecuenciaPago IS DISTINCT FROM @Frecuencia OR trabajador.FechaIngreso > @Hasta)
                THROW 51000, N'El personal del periodo cambio. Vuelva a consultarlo.', 1;

            INSERT INTO dbo.PlanPagoHYT (Frecuencia, Desde, Hasta) VALUES (@Frecuencia, @Desde, @Hasta);
            SET @PlanId = SCOPE_IDENTITY();
            INSERT INTO dbo.PlanPagoDetalleHYT (PlanId, TrabajadorId, Nombre, Tipo, Base, MinutosExtra, ImporteExtra, Descuento, Ajustes)
            SELECT @PlanId, fila.TrabajadorId, trabajador.NombreCompleto, trabajador.Tipo,
                   fila.Base, fila.MinutosExtra, fila.ImporteExtra, fila.Descuento, fila.Ajustes
            FROM @Datos AS fila INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = fila.TrabajadorId;
        END
        ELSE
        BEGIN
            IF EXISTS (SELECT fila.TrabajadorId FROM @Datos AS fila EXCEPT
                       SELECT detalle.TrabajadorId FROM dbo.PlanPagoDetalleHYT AS detalle WHERE detalle.PlanId = @PlanId)
               OR EXISTS (SELECT detalle.TrabajadorId FROM dbo.PlanPagoDetalleHYT AS detalle WHERE detalle.PlanId = @PlanId EXCEPT
                          SELECT fila.TrabajadorId FROM @Datos AS fila)
                THROW 51000, N'El detalle no coincide con la planificacion guardada.', 1;

            UPDATE detalle SET Base = fila.Base, MinutosExtra = fila.MinutosExtra,
                ImporteExtra = fila.ImporteExtra, Descuento = fila.Descuento, Ajustes = fila.Ajustes
            FROM dbo.PlanPagoDetalleHYT AS detalle INNER JOIN @Datos AS fila ON fila.TrabajadorId = detalle.TrabajadorId
            WHERE detalle.PlanId = @PlanId;
            UPDATE dbo.PlanPagoHYT SET Hasta = @Hasta WHERE PlanId = @PlanId;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PlanId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PlanPagoHYT', @RegistroId, N'GUARDAR';
        COMMIT TRANSACTION;
        SELECT @PlanId AS PlanId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
