-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza un asiento manual por empresa validando cuadre, periodo y correlativo mensual.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarAsientoManual
    @IdAsiento INT = NULL,
    @IdEmpresa INT,
    @IdOrigen INT,
    @FechaAsiento DATE,
    @Glosa NVARCHAR(500),
    @IdMoneda INT,
    @TipoCambio DECIMAL(18,6),
    @ReferenciaExterna NVARCHAR(100) = NULL,
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Ejercicio SMALLINT = YEAR(@FechaAsiento)
        DECLARE @Mes TINYINT = MONTH(@FechaAsiento)
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), YEAR(@FechaAsiento)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(@FechaAsiento)), 2)
        DECLARE @NumeroAsiento INT
        DECLARE @TotalDebe DECIMAL(18,2)
        DECLARE @TotalHaber DECIMAL(18,2)
        DECLARE @IdAsientoTrabajo INT
        DECLARE @PeriodoExistente CHAR(6)
        DECLARE @IdOrigenExistente INT

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle del asiento.', 16, 1);
        END;

        DECLARE @Detalle TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            ReferenciaLinea NVARCHAR(100) NULL
        );

        INSERT INTO @Detalle
        (
            Item,
            IdPlanCuenta,
            GlosaDetalle,
            Debe,
            Haber,
            ReferenciaLinea
        )
        SELECT
            T.N.value('@Item', 'smallint'),
            T.N.value('@IdPlanCuenta', 'int'),
            NULLIF(T.N.value('@GlosaDetalle', 'nvarchar(300)'), N''),
            T.N.value('@Debe', 'decimal(18,2)'),
            T.N.value('@Haber', 'decimal(18,2)'),
            NULLIF(T.N.value('@ReferenciaLinea', 'nvarchar(100)'), N'')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos una linea en el asiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            WHERE d.Item < 1
               OR d.Debe < 0
               OR d.Haber < 0
               OR ((d.Debe > 0 AND d.Haber > 0) OR (d.Debe = 0 AND d.Haber = 0))
        )
        BEGIN
            RAISERROR(N'Cada linea del asiento debe tener item valido y monto solo en Debe o Haber.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.Item
            FROM @Detalle AS d
            GROUP BY
                d.Item
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten items duplicados en el detalle del asiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            LEFT JOIN dbo.CON_PlanCuenta AS p
                ON p.IdPlanCuenta = d.IdPlanCuenta
               AND p.IdEmpresa = @IdEmpresa
               AND p.Estado = 1
               AND p.AceptaMovimiento = 1
            WHERE p.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR(N'Todas las cuentas del detalle deben existir, pertenecer a la empresa y aceptar movimiento.', 16, 1);
        END;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @Detalle AS d;

        IF ISNULL(@TotalDebe, 0) <= 0 OR ISNULL(@TotalHaber, 0) <= 0
        BEGIN
            RAISERROR(N'El asiento debe tener importes positivos tanto en Debe como en Haber.', 16, 1);
        END;

        IF @TotalDebe <> @TotalHaber
        BEGIN
            RAISERROR(N'El asiento no esta cuadrado. Debe y Haber deben ser iguales.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdOrigen = @IdOrigen
              AND o.IdEmpresa = @IdEmpresa
              AND o.Estado = 1
              AND o.PermiteRegistroManual = 1
        )
        BEGIN
            RAISERROR(N'El origen seleccionado no pertenece a la empresa o no permite registro manual.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Moneda AS m
            WHERE m.IdMoneda = @IdMoneda
              AND m.Estado = 1
        )
        BEGIN
            RAISERROR(N'La moneda seleccionada no esta activa.', 16, 1);
        END;

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

        BEGIN TRAN;

        IF @IdAsiento IS NULL
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @Periodo
            )
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = UltimoNumero + 1,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @Periodo;

                SELECT
                    @NumeroAsiento = c.UltimoNumero
                FROM dbo.CON_CorrelativoAsiento AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @Periodo;
            END
            ELSE
            BEGIN
                INSERT INTO dbo.CON_CorrelativoAsiento
                (
                    IdEmpresa,
                    IdOrigen,
                    Periodo,
                    UltimoNumero,
                    FechaActualizacion,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdOrigen,
                    @Periodo,
                    1,
                    SYSDATETIME(),
                    @UsuarioRegistro
                );

                SET @NumeroAsiento = 1;
            END;

            INSERT INTO dbo.CON_Asiento
            (
                IdEmpresa,
                IdOrigen,
                Ejercicio,
                Mes,
                Periodo,
                NumeroAsiento,
                FechaAsiento,
                Glosa,
                IdMoneda,
                TipoCambio,
                TotalDebe,
                TotalHaber,
                Estado,
                ReferenciaExterna,
                Observacion,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdOrigen,
                @Ejercicio,
                @Mes,
                @Periodo,
                @NumeroAsiento,
                @FechaAsiento,
                @Glosa,
                @IdMoneda,
                @TipoCambio,
                @TotalDebe,
                @TotalHaber,
                N'BORRADOR',
                @ReferenciaExterna,
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdAsientoTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT
                @IdAsientoTrabajo = a.IdAsiento,
                @NumeroAsiento = a.NumeroAsiento,
                @PeriodoExistente = a.Periodo,
                @IdOrigenExistente = a.IdOrigen
            FROM dbo.CON_Asiento AS a
            WHERE a.IdAsiento = @IdAsiento
              AND a.IdEmpresa = @IdEmpresa;

            IF @IdAsientoTrabajo IS NULL
            BEGIN
                RAISERROR(N'El asiento indicado no existe para la empresa activa.', 16, 1);
            END;

            IF @PeriodoExistente <> @Periodo
            BEGIN
                RAISERROR(N'No se puede cambiar el periodo del asiento existente. Mantenga la fecha dentro del mismo mes.', 16, 1);
            END;

            IF @IdOrigenExistente <> @IdOrigen
            BEGIN
                RAISERROR(N'No se puede cambiar el origen del asiento existente.', 16, 1);
            END;

            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.CON_Asiento
            SET FechaAsiento = @FechaAsiento,
                Glosa = @Glosa,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebe,
                TotalHaber = @TotalHaber,
                ReferenciaExterna = @ReferenciaExterna,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdAsiento = @IdAsientoTrabajo;
        END;

        INSERT INTO dbo.CON_AsientoDetalle
        (
            IdAsiento,
            Item,
            IdPlanCuenta,
            GlosaDetalle,
            Debe,
            Haber,
            ReferenciaLinea,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.GlosaDetalle,
            d.Debe,
            d.Haber,
            d.ReferenciaLinea,
            @UsuarioRegistro
        FROM @Detalle AS d
        ORDER BY
            d.Item ASC;

        COMMIT;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        SELECT
            a.IdAsiento,
            a.Periodo,
            a.NumeroAsiento,
            a.TotalDebe,
            a.TotalHaber,
            a.Estado
        FROM dbo.CON_Asiento AS a
        WHERE a.IdAsiento = @IdAsientoTrabajo;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
