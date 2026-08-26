-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Genera o regenera el asiento de cierre anual.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Crea un proceso anual CIE con opcion de regenerar cierres 14 y/o 15, tomando cuentas por ColBalance y usando el TC editable de 31/12.
-- Firma: FRANCO LARA - 03/07/2026 | Usa DH como sentido explicito para el calculo de saldos base y lo persiste en las dos lineas de cada asiento de cierre.
-- Firma: FRANCO LARA - 13/08/2026 | Reemplaza el cierre por cuenta por un unico asiento compuesto: acumula todas las cuentas con saldo hasta el periodo elegido, invierte su Debe/Haber, conserva importes reales en soles y dolares y genera el asiento en un periodo posterior seleccionado.
-- Firma: FRANCO LARA - 22/08/2026 | Limita el asiento compuesto a cuentas de Inventario (ColBalance = 'I'), no agrega cuentas de cuadre, permite Debe/Haber diferentes, corta como maximo en 13 y genera obligatoriamente en el periodo 14.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarCierreProceso
    @IdEmpresa INT,
    @Anio SMALLINT,
    @MesSaldoHasta TINYINT,
    @MesGeneracion TINYINT,
    @TipoCambioCompra DECIMAL(18,6),
    @TipoCambioVenta DECIMAL(18,6),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @PeriodoBaseInicio CHAR(6) = CONCAT(@Anio, '00')
        DECLARE @PeriodoSaldoHasta CHAR(6) = CONCAT(@Anio, RIGHT('0' + CONVERT(VARCHAR(2), @MesSaldoHasta), 2))
        DECLARE @PeriodoGeneracion CHAR(6) = CONCAT(@Anio, RIGHT('0' + CONVERT(VARCHAR(2), @MesGeneracion), 2))
        DECLARE @FechaAsiento DATE = DATEFROMPARTS(@Anio, 12, 31)
        DECLARE @IdOrigen INT
        DECLARE @UsaTipoCambioSbs BIT = 0
        DECLARE @IdCuentaAdministradora INT
        DECLARE @IdMonedaPen INT
        DECLARE @IdCierreProceso INT
        DECLARE @IdAsientoTrabajo INT
        DECLARE @NumeroAsientoTrabajo INT
        DECLARE @TotalLineas INT = 0
        DECLARE @TotalDebeProceso DECIMAL(18,2) = 0
        DECLARE @TotalHaberProceso DECIMAL(18,2) = 0

        IF @Anio < 2000 OR @Anio > 9999
        BEGIN
            RAISERROR(N'El ejercicio indicado es invalido.', 16, 1);
        END;

        IF @MesSaldoHasta > 13
        BEGIN
            RAISERROR(N'El periodo contable de corte debe estar entre 00 y 13.', 16, 1);
        END;

        IF @MesGeneracion <> 14
        BEGIN
            RAISERROR(N'El asiento de cierre de Inventario debe generarse en el periodo 14.', 16, 1);
        END;

        IF @MesGeneracion <= @MesSaldoHasta
        BEGIN
            RAISERROR(N'El periodo de generacion debe ser posterior al periodo usado como corte.', 16, 1);
        END;

        IF ISNULL(@TipoCambioCompra, 0) <= 0 OR ISNULL(@TipoCambioVenta, 0) <= 0
        BEGIN
            RAISERROR(N'Ingrese un tipo de cambio de cierre mayor a cero.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'CIE'
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.Activo = 1;

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'No existe una configuracion activa de asiento de cierre en configuracion contable.', 16, 1);
        END;

        SELECT
            @UsaTipoCambioSbs = CASE
                                    WHEN UPPER(LTRIM(RTRIM(pe.ValorParametro))) = 'S' THEN 1
                                    ELSE 0
                                END
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'TIPO_CAMBIO_SBS_CIERRE'
          AND pe.Activo = 1;

        SELECT
            @IdCuentaAdministradora = e.IdCuentaAdministradora
        FROM dbo.SEG_Empresa AS e
        WHERE e.IdEmpresa = @IdEmpresa;

        IF @IdCuentaAdministradora IS NULL
        BEGIN
            RAISERROR(N'La empresa no tiene una cuenta administradora asociada para obtener el tipo de cambio.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_TipoCambio AS tc
            WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
              AND tc.Fecha = @FechaAsiento
              AND tc.IdMoneda = 'USD'
              AND tc.Estado = 1
        )
        BEGIN
            RAISERROR(N'No existe tipo de cambio USD registrado para el 31/12 del ejercicio seleccionado.', 16, 1);
        END;

        SELECT
            @IdMonedaPen = m.IdMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.CodigoMoneda = 'PEN'
          AND m.Estado = 1;

        IF @IdMonedaPen IS NULL
        BEGIN
            RAISERROR(N'La moneda PEN no esta registrada como activa.', 16, 1);
        END;

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        DECLARE @CorrelativosRecalcular TABLE
        (
            IdOrigen INT NOT NULL,
            Periodo CHAR(6) NOT NULL,
            PRIMARY KEY (IdOrigen, Periodo)
        );

        INSERT INTO @CorrelativosRecalcular (IdOrigen, Periodo)
        VALUES (@IdOrigen, @PeriodoGeneracion);

        DECLARE @LineasGeneradas TABLE
        (
            OrdenLinea INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            IdPlanCuenta INT NOT NULL,
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            CodigoMoneda VARCHAR(3) NOT NULL,
            TipoCambioAplicado DECIMAL(18,6) NOT NULL,
            DH CHAR(1) NOT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            TotalImporteS DECIMAL(18,2) NOT NULL,
            TotalImporteD DECIMAL(18,2) NOT NULL
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        UPDATE dbo.CON_TipoCambio
        SET Compra = CASE WHEN @UsaTipoCambioSbs = 1 THEN Compra ELSE @TipoCambioCompra END,
            Venta = CASE WHEN @UsaTipoCambioSbs = 1 THEN Venta ELSE @TipoCambioVenta END,
            CompraSBS = CASE WHEN @UsaTipoCambioSbs = 1 THEN @TipoCambioCompra ELSE CompraSBS END,
            VentaSBS = CASE WHEN @UsaTipoCambioSbs = 1 THEN @TipoCambioVenta ELSE VentaSBS END,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdCuentaAdministradora = @IdCuentaAdministradora
          AND Fecha = @FechaAsiento
          AND IdMoneda = 'USD'
          AND Estado = 1;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_CierreProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio
        )
        BEGIN
            INSERT INTO @AsientosEliminar (IdAsiento)
            SELECT DISTINCT
                x.IdAsiento
            FROM
            (
                SELECT p.IdAsiento
                FROM dbo.CON_CierreProceso AS p
                WHERE p.IdEmpresa = @IdEmpresa
                  AND p.Anio = @Anio
                  AND p.IdAsiento IS NOT NULL

                UNION

                SELECT d.IdAsiento
                FROM dbo.CON_CierreProcesoDetalle AS d
                INNER JOIN dbo.CON_CierreProceso AS p
                    ON p.IdCierreProceso = d.IdCierreProceso
                WHERE p.IdEmpresa = @IdEmpresa
                  AND p.Anio = @Anio
                  AND d.IdAsiento IS NOT NULL
            ) AS x;

            INSERT INTO @CorrelativosRecalcular (IdOrigen, Periodo)
            SELECT DISTINCT
                a.IdOrigen,
                a.Periodo
            FROM dbo.CON_Asiento AS a
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = a.IdAsiento
            WHERE NOT EXISTS
            (
                SELECT 1
                FROM @CorrelativosRecalcular AS r
                WHERE r.IdOrigen = a.IdOrigen
                  AND r.Periodo = a.Periodo
            );

            DELETE d
            FROM dbo.CON_CierreProcesoDetalle AS d
            INNER JOIN dbo.CON_CierreProceso AS p
                ON p.IdCierreProceso = d.IdCierreProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio;

            DELETE p
            FROM dbo.CON_CierreProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio;

            DELETE d
            FROM dbo.CON_AsientoDetalle AS d
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = d.IdAsiento;

            DELETE a
            FROM dbo.CON_Asiento AS a
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = a.IdAsiento;
        END;

        UPDATE correlativo
        SET UltimoNumero = base.UltimoNumero,
            FechaActualizacion = SYSDATETIME(),
            UsuarioRegistro = @UsuarioRegistro
        FROM dbo.CON_CorrelativoAsiento AS correlativo
        INNER JOIN @CorrelativosRecalcular AS r
            ON r.IdOrigen = correlativo.IdOrigen
           AND r.Periodo = correlativo.Periodo
        INNER JOIN
        (
            SELECT
                a.IdOrigen,
                a.Periodo,
                MAX(a.NumeroAsiento) AS UltimoNumero
            FROM dbo.CON_Asiento AS a
            INNER JOIN @CorrelativosRecalcular AS r
                ON r.IdOrigen = a.IdOrigen
               AND r.Periodo = a.Periodo
            WHERE a.IdEmpresa = @IdEmpresa
            GROUP BY
                a.IdOrigen,
                a.Periodo
        ) AS base
            ON base.IdOrigen = correlativo.IdOrigen
           AND base.Periodo = correlativo.Periodo
        WHERE correlativo.IdEmpresa = @IdEmpresa;

        DELETE correlativo
        FROM dbo.CON_CorrelativoAsiento AS correlativo
        INNER JOIN @CorrelativosRecalcular AS r
            ON r.IdOrigen = correlativo.IdOrigen
           AND r.Periodo = correlativo.Periodo
        WHERE correlativo.IdEmpresa = @IdEmpresa
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.CON_Asiento AS a
              WHERE a.IdEmpresa = @IdEmpresa
                AND a.IdOrigen = correlativo.IdOrigen
                AND a.Periodo = correlativo.Periodo
          );

        ;WITH SaldosCuenta AS
        (
            SELECT
                d.IdPlanCuenta,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                CASE WHEN NULLIF(pc.IdMoneda, '') = 'USD' THEN 'USD' ELSE 'PEN' END AS CodigoMoneda,
                ISNULL(NULLIF(pc.TipoCambio, ''), 'V') AS TipoCambioCuenta,
                CAST(SUM(
                    CASE
                        WHEN d.DH = 'D' THEN CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Debe END
                        ELSE CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Haber END * -1
                    END
                ) AS DECIMAL(18,2)) AS SaldoSoles,
                CAST(SUM(
                    CASE
                        WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteD, 0)
                        ELSE ISNULL(d.TotalImporteD, 0) * -1
                    END
                ) AS DECIMAL(18,2)) AS SaldoDolares
            FROM dbo.CON_Asiento AS a
            INNER JOIN dbo.CON_AsientoDetalle AS d
                ON d.IdAsiento = a.IdAsiento
            INNER JOIN dbo.CON_PlanCuenta AS pc
                ON pc.IdPlanCuenta = d.IdPlanCuenta
               AND pc.IdEmpresa = @IdEmpresa
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.Periodo BETWEEN @PeriodoBaseInicio AND @PeriodoSaldoHasta
              AND pc.ColBalance = 'I'
            GROUP BY
                d.IdPlanCuenta,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                pc.IdMoneda,
                pc.TipoCambio
            HAVING ABS(SUM(
                CASE
                    WHEN d.DH = 'D' THEN CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Debe END
                    ELSE CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Haber END * -1
                END
            )) >= 0.01
        )
        INSERT INTO @LineasGeneradas
        (
            IdPlanCuenta,
            CodigoCuenta,
            NombreCuenta,
            CodigoMoneda,
            TipoCambioAplicado,
            DH,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD
        )
        SELECT
            s.IdPlanCuenta,
            s.CodigoCuenta,
            s.NombreCuenta,
            s.CodigoMoneda,
            CASE
                WHEN s.CodigoMoneda = 'USD' AND s.TipoCambioCuenta = 'C' THEN @TipoCambioCompra
                WHEN s.CodigoMoneda = 'USD' THEN @TipoCambioVenta
                ELSE 1
            END,
            CASE WHEN s.SaldoSoles > 0 THEN 'H' ELSE 'D' END,
            CASE WHEN s.SaldoSoles < 0 THEN ABS(s.SaldoSoles) ELSE 0 END,
            CASE WHEN s.SaldoSoles > 0 THEN ABS(s.SaldoSoles) ELSE 0 END,
            ABS(s.SaldoSoles),
            ABS(s.SaldoDolares)
        FROM SaldosCuenta AS s
        ORDER BY s.CodigoCuenta;

        SELECT
            @TotalLineas = COUNT(*),
            @TotalDebeProceso = ISNULL(SUM(l.Debe), 0),
            @TotalHaberProceso = ISNULL(SUM(l.Haber), 0)
        FROM @LineasGeneradas AS l;

        IF @TotalLineas = 0
        BEGIN
            RAISERROR(N'No existen cuentas configuradas como Inventario con saldo pendiente para generar el asiento de cierre.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdOrigen = @IdOrigen
              AND c.Periodo = @PeriodoGeneracion
        )
        BEGIN
            UPDATE dbo.CON_CorrelativoAsiento
            SET UltimoNumero = UltimoNumero + 1,
                FechaActualizacion = SYSDATETIME(),
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdOrigen = @IdOrigen
              AND Periodo = @PeriodoGeneracion;

            SELECT
                @NumeroAsientoTrabajo = c.UltimoNumero
            FROM dbo.CON_CorrelativoAsiento AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdOrigen = @IdOrigen
              AND c.Periodo = @PeriodoGeneracion;
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
                @PeriodoGeneracion,
                1,
                SYSDATETIME(),
                @UsuarioRegistro
            );

            SET @NumeroAsientoTrabajo = 1;
        END;

        INSERT INTO dbo.CON_Asiento
        (
            IdEmpresa,
            IdOrigen,
            Ejercicio,
            Mes,
            Periodo,
            NumeroAsiento,
            FechaEmision,
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
            @Anio,
            @MesGeneracion,
            @PeriodoGeneracion,
            @NumeroAsientoTrabajo,
            @FechaAsiento,
            @FechaAsiento,
            N'ASIENTO DE CIERRE',
            @IdMonedaPen,
            @TipoCambioVenta,
            @TotalDebeProceso,
            @TotalHaberProceso,
            N'PROVISIONADO',
            CONCAT(N'CIE-', @Anio),
            CONCAT(N'Generado automaticamente con saldos hasta ', @PeriodoSaldoHasta, N' en el periodo ', @PeriodoGeneracion, N'.'),
            @UsuarioRegistro
        );

        SET @IdAsientoTrabajo = SCOPE_IDENTITY();

        INSERT INTO dbo.CON_AsientoDetalle
        (
            IdAsiento,
            Item,
            IdPlanCuenta,
            DH,
            GlosaDetalle,
            TipoCambioLinea,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            ROW_NUMBER() OVER (ORDER BY l.CodigoCuenta, l.OrdenLinea),
            l.IdPlanCuenta,
            l.DH,
            CONCAT(N'CIERRE / ', l.CodigoCuenta, N' - ', l.NombreCuenta),
            l.TipoCambioAplicado,
            l.Debe,
            l.Haber,
            l.TotalImporteS,
            l.TotalImporteD,
            @UsuarioRegistro
        FROM @LineasGeneradas AS l;

        INSERT INTO dbo.CON_CierreProceso
        (
            IdEmpresa,
            Anio,
            MesSaldoHasta,
            PeriodoSaldoHasta,
            MesGeneracion,
            PeriodoGeneracion,
            IdOrigen,
            FechaAsiento,
            UsaTipoCambioSbs,
            TipoCambioCompra,
            TipoCambioVenta,
            ProcesaGananciasPerdidas,
            ProcesaInventarios,
            IdAsiento,
            NumeroAsiento,
            TotalLineas,
            TotalCuentas,
            TotalAsientos,
            TotalDebe,
            TotalHaber,
            UsuarioRegistro
        )
        VALUES
        (
            @IdEmpresa,
            @Anio,
            @MesSaldoHasta,
            @PeriodoSaldoHasta,
            @MesGeneracion,
            @PeriodoGeneracion,
            @IdOrigen,
            @FechaAsiento,
            @UsaTipoCambioSbs,
            @TipoCambioCompra,
            @TipoCambioVenta,
            0,
            1,
            @IdAsientoTrabajo,
            @NumeroAsientoTrabajo,
            @TotalLineas,
            @TotalLineas,
            1,
            @TotalDebeProceso,
            @TotalHaberProceso,
            @UsuarioRegistro
        );

        SET @IdCierreProceso = SCOPE_IDENTITY();

        INSERT INTO dbo.CON_CierreProcesoDetalle
        (
            IdCierreProceso,
            Item,
            TipoCierre,
            IdPlanCuenta,
            CodigoMoneda,
            TipoCambioAplicado,
            IdAsiento,
            NumeroAsiento,
            DH,
            TotalDebe,
            TotalHaber,
            TotalImporteS,
            TotalImporteD,
            Estado,
            Observacion,
            UsuarioRegistro
        )
        SELECT
            @IdCierreProceso,
            ROW_NUMBER() OVER (ORDER BY l.CodigoCuenta, l.OrdenLinea),
            RIGHT(@PeriodoGeneracion, 2),
            l.IdPlanCuenta,
            l.CodigoMoneda,
            l.TipoCambioAplicado,
            @IdAsientoTrabajo,
            @NumeroAsientoTrabajo,
            l.DH,
            l.Debe,
            l.Haber,
            l.TotalImporteS,
            l.TotalImporteD,
            N'GENERADO',
            CONCAT(N'Saldo acumulado invertido hasta el periodo ', @PeriodoSaldoHasta, N'.'),
            @UsuarioRegistro
        FROM @LineasGeneradas AS l;

        COMMIT;
        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        EXEC dbo.usp_CON_ObtenerCierreProceso
            @IdEmpresa = @IdEmpresa,
            @Anio = @Anio;

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
