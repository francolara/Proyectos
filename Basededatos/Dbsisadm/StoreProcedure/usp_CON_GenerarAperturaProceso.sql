-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Genera o regenera el asiento de apertura anual en el periodo 00.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Genera un unico asiento de apertura por ejercicio usando saldos acumulados hasta un periodo de corte del anio anterior y la logica referencial equivalente a usp_AsientodeApertura_2, agrupando el analisis por numero de documento, tipo, serie y referencia sin heredar cliente/proveedor al asiento generado.
-- Firma: FRANCO LARA - 03/07/2026 | Toma DH como marca explicita del sentido contable al calcular los saldos base y lo persiste en el detalle del asiento de apertura.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarAperturaProceso
    @IdEmpresa INT,
    @AnioApertura SMALLINT,
    @MesSaldoHasta TINYINT,
    @TipoCambioCompra DECIMAL(18,6),
    @TipoCambioVenta DECIMAL(18,6),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @AnioSaldo SMALLINT = @AnioApertura - 1
        DECLARE @PeriodoApertura CHAR(6) = CONCAT(@AnioApertura, '00')
        DECLARE @PeriodoSaldoHasta CHAR(6) = CONCAT(@AnioSaldo, RIGHT('0' + CONVERT(VARCHAR(2), @MesSaldoHasta), 2))
        DECLARE @FechaAsiento DATE = DATEFROMPARTS(@AnioApertura, 1, 1)
        DECLARE @FechaTipoCambio DATE = DATEFROMPARTS(@AnioSaldo, 12, 31)
        DECLARE @IdOrigen INT
        DECLARE @UsaTipoCambioSbs BIT = 0
        DECLARE @IdMonedaPen INT
        DECLARE @IdAperturaProceso INT
        DECLARE @NumeroAsientoTrabajo INT = NULL
        DECLARE @IdAsientoTrabajo INT = NULL
        DECLARE @TotalLineas INT = 0
        DECLARE @TotalDebeProceso DECIMAL(18,2) = 0
        DECLARE @TotalHaberProceso DECIMAL(18,2) = 0
        DECLARE @UltimoNumeroRestante INT = 0

        IF @AnioApertura < 2000 OR @AnioApertura > 9999
        BEGIN
            RAISERROR(N'El anio de apertura es invalido.', 16, 1);
        END;

        IF @MesSaldoHasta > 15
        BEGIN
            RAISERROR(N'El mes contable de corte es invalido.', 16, 1);
        END;

        IF ISNULL(@TipoCambioCompra, 0) <= 0 OR ISNULL(@TipoCambioVenta, 0) <= 0
        BEGIN
            RAISERROR(N'Ingrese un tipo de cambio de apertura mayor a cero.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'APR'
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.Activo = 1;

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'No existe una configuracion activa de asiento de apertura en configuracion contable.', 16, 1);
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

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_TipoCambio AS tc
            INNER JOIN dbo.SEG_Empresa AS e
                ON e.IdCuentaAdministradora = tc.IdCuentaAdministradora
            WHERE e.IdEmpresa = @IdEmpresa
              AND tc.Fecha = @FechaTipoCambio
              AND tc.IdMoneda = 'USD'
              AND tc.Estado = 1
        )
        BEGIN
            RAISERROR(N'No existe tipo de cambio USD registrado para el 31/12 del anio base de apertura.', 16, 1);
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

        DECLARE @LineasGeneradas TABLE
        (
            OrdenLinea INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            TipoDetalle NVARCHAR(20) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NOT NULL,
            CodigoMoneda VARCHAR(3) NOT NULL,
            TipoCambioAplicado DECIMAL(18,6) NOT NULL,
            NumeroDocumento VARCHAR(20) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            IdCliente INT NULL,
            IdProveedor INT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            TotalImporteS DECIMAL(18,2) NOT NULL,
            TotalImporteD DECIMAL(18,2) NOT NULL,
            Observacion NVARCHAR(300) NULL
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_AperturaProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.AnioApertura = @AnioApertura
        )
        BEGIN
            INSERT INTO @AsientosEliminar (IdAsiento)
            SELECT DISTINCT
                p.IdAsiento
            FROM dbo.CON_AperturaProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.AnioApertura = @AnioApertura
              AND p.IdAsiento IS NOT NULL;

            DELETE d
            FROM dbo.CON_AperturaProcesoDetalle AS d
            INNER JOIN dbo.CON_AperturaProceso AS p
                ON p.IdAperturaProceso = d.IdAperturaProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.AnioApertura = @AnioApertura;

            DELETE p
            FROM dbo.CON_AperturaProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.AnioApertura = @AnioApertura;

            DELETE d
            FROM dbo.CON_AsientoDetalle AS d
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = d.IdAsiento;

            DELETE a
            FROM dbo.CON_Asiento AS a
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = a.IdAsiento;

            SELECT
                @UltimoNumeroRestante = ISNULL(MAX(a.NumeroAsiento), 0)
            FROM dbo.CON_Asiento AS a
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.IdOrigen = @IdOrigen
              AND a.Periodo = @PeriodoApertura;

            IF @UltimoNumeroRestante = 0
            BEGIN
                DELETE correlativo
                FROM dbo.CON_CorrelativoAsiento AS correlativo
                WHERE correlativo.IdEmpresa = @IdEmpresa
                  AND correlativo.IdOrigen = @IdOrigen
                  AND correlativo.Periodo = @PeriodoApertura;
            END
            ELSE
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = @UltimoNumeroRestante,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @PeriodoApertura;
            END;
        END;

        ;WITH MovimientosBase AS
        (
            SELECT
                d.IdPlanCuenta,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                CASE WHEN NULLIF(pc.IdMoneda, '') = 'USD' THEN 'USD' ELSE 'PEN' END AS CodigoMoneda,
                CASE WHEN LEFT(pc.CodigoCuenta, 1) < '4' THEN @TipoCambioCompra ELSE @TipoCambioVenta END AS TipoCambioAplicado,
                NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), '') AS NumeroDocumento,
                NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N'') AS TipoDocumento,
                NULLIF(LTRIM(RTRIM(d.Serie)), '') AS Serie,
                NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), '') AS ReferenciaLinea,
                CASE
                    WHEN d.DH = 'D' THEN CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Debe END
                    ELSE CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Haber END * -1
                END AS MovimientoSoles,
                CASE
                    WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteD, 0)
                    ELSE ISNULL(d.TotalImporteD, 0) * -1
                END AS MovimientoDolares
            FROM dbo.CON_Asiento AS a
            INNER JOIN dbo.CON_AsientoDetalle AS d
                ON d.IdAsiento = a.IdAsiento
            INNER JOIN dbo.CON_PlanCuenta AS pc
                ON pc.IdPlanCuenta = d.IdPlanCuenta
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.Periodo BETWEEN CONCAT(@AnioSaldo, '00') AND @PeriodoSaldoHasta
              AND pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND pc.AceptaMovimiento = 1
              AND LEFT(pc.CodigoCuenta, 1) BETWEEN '1' AND '5'
        ),
        ResumenCuenta AS
        (
            SELECT
                N'RESUMEN' AS TipoDetalle,
                m.IdPlanCuenta,
                MAX(m.NombreCuenta) AS NombreCuenta,
                MAX(m.CodigoMoneda) AS CodigoMoneda,
                MAX(m.TipoCambioAplicado) AS TipoCambioAplicado,
                CAST(NULL AS VARCHAR(20)) AS NumeroDocumento,
                CAST(NULL AS NVARCHAR(150)) AS TipoDocumento,
                CAST(NULL AS VARCHAR(10)) AS Serie,
                CAST(NULL AS NVARCHAR(100)) AS ReferenciaLinea,
                SUM(m.MovimientoSoles) AS SaldoSoles,
                SUM(m.MovimientoDolares) AS SaldoDolares,
                N'Saldo consolidado sin referencia documental.' AS Observacion
            FROM MovimientosBase AS m
            WHERE m.NumeroDocumento IS NULL
              AND m.TipoDocumento IS NULL
              AND m.Serie IS NULL
              AND m.ReferenciaLinea IS NULL
            GROUP BY
                m.IdPlanCuenta
            HAVING ABS(SUM(m.MovimientoSoles)) > 0.004
        ),
        AnalisisCuenta AS
        (
            SELECT
                N'ANALISIS' AS TipoDetalle,
                m.IdPlanCuenta,
                MAX(m.NombreCuenta) AS NombreCuenta,
                MAX(m.CodigoMoneda) AS CodigoMoneda,
                MAX(m.TipoCambioAplicado) AS TipoCambioAplicado,
                m.NumeroDocumento,
                m.TipoDocumento,
                m.Serie,
                m.ReferenciaLinea,
                SUM(m.MovimientoSoles) AS SaldoSoles,
                SUM(m.MovimientoDolares) AS SaldoDolares,
                N'Saldo agrupado por referencia documental o auxiliar.' AS Observacion
            FROM MovimientosBase AS m
            WHERE m.NumeroDocumento IS NOT NULL
               OR m.TipoDocumento IS NOT NULL
               OR m.Serie IS NOT NULL
               OR m.ReferenciaLinea IS NOT NULL
            GROUP BY
                m.IdPlanCuenta,
                m.NumeroDocumento,
                m.TipoDocumento,
                m.Serie,
                m.ReferenciaLinea
            HAVING ABS(SUM(m.MovimientoSoles)) > 0.004
        )
        INSERT INTO @LineasGeneradas
        (
            TipoDetalle,
            IdPlanCuenta,
            GlosaDetalle,
            CodigoMoneda,
            TipoCambioAplicado,
            NumeroDocumento,
            TipoDocumento,
            Serie,
            ReferenciaLinea,
            IdCliente,
            IdProveedor,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            Observacion
        )
        SELECT
            origen.TipoDetalle,
            origen.IdPlanCuenta,
            LEFT(
                CONCAT(
                    N'ASIENTO DE APERTURA ',
                    pc.CodigoCuenta,
                    N' - ',
                    pc.NombreCuenta,
                    CASE
                        WHEN origen.TipoDetalle = N'ANALISIS'
                             AND origen.NumeroDocumento IS NOT NULL
                            THEN CONCAT(N' / ', ISNULL(origen.TipoDocumento, N''), N' ', ISNULL(origen.Serie, ''), N'-', origen.NumeroDocumento, CASE WHEN origen.ReferenciaLinea IS NOT NULL THEN CONCAT(N' / ', origen.ReferenciaLinea) ELSE N'' END)
                        ELSE N''
                    END
                ),
                300
            ),
            origen.CodigoMoneda,
            origen.TipoCambioAplicado,
            origen.NumeroDocumento,
            origen.TipoDocumento,
            origen.Serie,
            origen.ReferenciaLinea,
            NULL,
            NULL,
            CASE WHEN origen.SaldoSoles > 0 THEN ROUND(origen.SaldoSoles, 2) ELSE 0 END,
            CASE WHEN origen.SaldoSoles < 0 THEN ROUND(ABS(origen.SaldoSoles), 2) ELSE 0 END,
            ROUND(ABS(origen.SaldoSoles), 2),
            ROUND(ABS(origen.SaldoDolares), 2),
            origen.Observacion
        FROM
        (
            SELECT * FROM ResumenCuenta
            UNION ALL
            SELECT * FROM AnalisisCuenta
        ) AS origen
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = origen.IdPlanCuenta
        ORDER BY
            pc.CodigoCuenta ASC,
            CASE origen.TipoDetalle WHEN N'RESUMEN' THEN 0 ELSE 1 END,
            ISNULL(origen.TipoDocumento, N'') ASC,
            ISNULL(origen.Serie, '') ASC,
            ISNULL(origen.NumeroDocumento, '') ASC;

        SET @TotalLineas = (SELECT COUNT(*) FROM @LineasGeneradas);
        SET @TotalDebeProceso = ISNULL((SELECT SUM(Debe) FROM @LineasGeneradas), 0);
        SET @TotalHaberProceso = ISNULL((SELECT SUM(Haber) FROM @LineasGeneradas), 0);

        IF @TotalLineas = 0
        BEGIN
            RAISERROR(N'No existen saldos pendientes para generar el asiento de apertura con el corte seleccionado.', 16, 1);
        END;

        IF ABS(@TotalDebeProceso - @TotalHaberProceso) > 0.01
        BEGIN
            RAISERROR(N'El asiento de apertura calculado no queda cuadrado con el corte seleccionado.', 16, 1);
        END;

        INSERT INTO dbo.CON_AperturaProceso
        (
            IdEmpresa,
            AnioApertura,
            AnioSaldo,
            MesSaldoHasta,
            PeriodoSaldoHasta,
            IdOrigen,
            FechaAsiento,
            UsaTipoCambioSbs,
            TipoCambioCompra,
            TipoCambioVenta,
            IdAsiento,
            NumeroAsiento,
            TotalLineas,
            TotalDebe,
            TotalHaber,
            UsuarioRegistro
        )
        VALUES
        (
            @IdEmpresa,
            @AnioApertura,
            @AnioSaldo,
            @MesSaldoHasta,
            @PeriodoSaldoHasta,
            @IdOrigen,
            @FechaAsiento,
            @UsaTipoCambioSbs,
            @TipoCambioCompra,
            @TipoCambioVenta,
            NULL,
            NULL,
            @TotalLineas,
            @TotalDebeProceso,
            @TotalHaberProceso,
            @UsuarioRegistro
        );

        SET @IdAperturaProceso = SCOPE_IDENTITY();

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdOrigen = @IdOrigen
              AND c.Periodo = @PeriodoApertura
        )
        BEGIN
            UPDATE dbo.CON_CorrelativoAsiento
            SET UltimoNumero = UltimoNumero + 1,
                FechaActualizacion = SYSDATETIME(),
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdOrigen = @IdOrigen
              AND Periodo = @PeriodoApertura;

            SELECT
                @NumeroAsientoTrabajo = c.UltimoNumero
            FROM dbo.CON_CorrelativoAsiento AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdOrigen = @IdOrigen
              AND c.Periodo = @PeriodoApertura;
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
                @PeriodoApertura,
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
            @AnioApertura,
            0,
            @PeriodoApertura,
            @NumeroAsientoTrabajo,
            @FechaAsiento,
            @FechaAsiento,
            N'ASIENTO DE APERTURA',
            @IdMonedaPen,
            @TipoCambioCompra,
            @TotalDebeProceso,
            @TotalHaberProceso,
            N'PROVISIONADO',
            CONCAT(N'APR-', @AnioApertura),
            CONCAT(N'Generado automaticamente con saldos hasta ', @PeriodoSaldoHasta, N'.'),
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
            NumeroDocumento,
            TipoDocumento,
            Serie,
            ReferenciaLinea,
            TipoCambioLinea,
            IdCliente,
            IdProveedor,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            ROW_NUMBER() OVER (
                ORDER BY
                    pc.CodigoCuenta ASC,
                    CASE l.TipoDetalle WHEN N'RESUMEN' THEN 0 ELSE 1 END,
                    ISNULL(l.TipoDocumento, N'') ASC,
                    ISNULL(l.Serie, '') ASC,
                    ISNULL(l.NumeroDocumento, '') ASC,
                    l.OrdenLinea ASC
            ),
            l.IdPlanCuenta,
            CASE WHEN l.Debe > 0 THEN 'D' ELSE 'H' END,
            l.GlosaDetalle,
            l.NumeroDocumento,
            l.TipoDocumento,
            l.Serie,
            l.ReferenciaLinea,
            l.TipoCambioAplicado,
            NULL,
            NULL,
            l.Debe,
            l.Haber,
            l.TotalImporteS,
            l.TotalImporteD,
            @UsuarioRegistro
        FROM @LineasGeneradas AS l
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = l.IdPlanCuenta;

        INSERT INTO dbo.CON_AperturaProcesoDetalle
        (
            IdAperturaProceso,
            Item,
            TipoDetalle,
            IdPlanCuenta,
            CodigoMoneda,
            TipoCambioAplicado,
            TipoDocumento,
            Serie,
            NumeroDocumento,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            Observacion,
            UsuarioRegistro
        )
        SELECT
            @IdAperturaProceso,
            ROW_NUMBER() OVER (
                ORDER BY
                    pc.CodigoCuenta ASC,
                    CASE l.TipoDetalle WHEN N'RESUMEN' THEN 0 ELSE 1 END,
                    ISNULL(l.TipoDocumento, N'') ASC,
                    ISNULL(l.Serie, '') ASC,
                    ISNULL(l.NumeroDocumento, '') ASC,
                    l.OrdenLinea ASC
            ),
            l.TipoDetalle,
            l.IdPlanCuenta,
            l.CodigoMoneda,
            l.TipoCambioAplicado,
            l.TipoDocumento,
            l.Serie,
            l.NumeroDocumento,
            l.Debe,
            l.Haber,
            l.TotalImporteS,
            l.TotalImporteD,
            l.Observacion,
            @UsuarioRegistro
        FROM @LineasGeneradas AS l
        INNER JOIN dbo.CON_PlanCuenta AS pc
            ON pc.IdPlanCuenta = l.IdPlanCuenta;

        UPDATE dbo.CON_AperturaProceso
        SET IdAsiento = @IdAsientoTrabajo,
            NumeroAsiento = @NumeroAsientoTrabajo,
            TotalLineas = @TotalLineas,
            TotalDebe = @TotalDebeProceso,
            TotalHaber = @TotalHaberProceso,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdAperturaProceso = @IdAperturaProceso;

        COMMIT;
        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        EXEC dbo.usp_CON_ObtenerAperturaProceso
            @IdEmpresa = @IdEmpresa,
            @AnioApertura = @AnioApertura;

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
