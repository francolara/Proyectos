-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Genera o regenera los asientos de ajuste de cuentas por cuenta analitica para un periodo.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Replica el ajuste de cuentas del legacy solo para cuentas de analisis, crea un asiento AJU separado por cuenta, usa CUENTAGANANCIA_AJ y CUENTAPERDIDA_AJ, vuelve a expandir cuentas destino/contrapartida cuando la regla exista, agrupa el analisis por auxiliar-documento usando NumeroDocumento como ctaauxiliar web, limpia las tablas variables por iteracion para evitar arrastre de cuentas entre asientos, no hereda cliente/proveedor al asiento generado, excluye asientos originados por procesos automaticos DIF/AJU/APR/CIE para no recalcular sobre ajustes ya generados y ahora genera cada asiento en la moneda natural de la cuenta, conservando los equivalentes en soles y dolares del detalle.
-- Firma: FRANCO LARA - 03/07/2026 | Usa DH para leer el sentido historico del detalle contable y lo persiste tambien en cada linea generada del ajuste.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarAjusteCuentaProceso
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Anio SMALLINT
        DECLARE @Mes TINYINT
        DECLARE @FechaAsiento DATE
        DECLARE @IdOrigen INT
        DECLARE @IdMonedaPen INT
        DECLARE @IdMonedaUsd INT
        DECLARE @IdCuentaAdministradora INT
        DECLARE @TipoCambioCompra DECIMAL(18,6) = 0
        DECLARE @TipoCambioVenta DECIMAL(18,6) = 0
        DECLARE @CodigoCuentaGanancia VARCHAR(20)
        DECLARE @CodigoCuentaPerdida VARCHAR(20)
        DECLARE @IdPlanCuentaGanancia INT
        DECLARE @IdPlanCuentaPerdida INT
        DECLARE @IdAjusteCuentaProceso INT
        DECLARE @TotalCuentas INT = 0
        DECLARE @TotalAsientos INT = 0
        DECLARE @TotalDebeProceso DECIMAL(18,2) = 0
        DECLARE @TotalHaberProceso DECIMAL(18,2) = 0

        IF @Periodo IS NULL
           OR @Periodo NOT LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
           OR RIGHT(@Periodo, 2) NOT BETWEEN '01' AND '12'
        BEGIN
            RAISERROR(N'El periodo debe estar en formato yyyyMM.', 16, 1);
        END;

        SET @Anio = TRY_CONVERT(SMALLINT, LEFT(@Periodo, 4));
        SET @Mes = TRY_CONVERT(TINYINT, RIGHT(@Periodo, 2));
        SET @FechaAsiento = EOMONTH(DATEFROMPARTS(@Anio, @Mes, 1));

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PeriodoContableEstado AS pe
            WHERE pe.IdEmpresa = @IdEmpresa
              AND pe.Periodo = @Periodo
              AND pe.Cerrado = 1
        )
        BEGIN
            RAISERROR(N'El periodo seleccionado se encuentra cerrado.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'AJU'
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.Activo = 1;

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'No existe una configuracion activa de ajuste de cuentas en configuracion contable.', 16, 1);
        END;

        SELECT
            @CodigoCuentaGanancia = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAGANANCIA_AJ'
          AND pe.Activo = 1;

        SELECT
            @CodigoCuentaPerdida = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAPERDIDA_AJ'
          AND pe.Activo = 1;

        IF @CodigoCuentaGanancia IS NULL OR @CodigoCuentaPerdida IS NULL
        BEGIN
            RAISERROR(N'Configure las cuentas CUENTAGANANCIA_AJ y CUENTAPERDIDA_AJ para la empresa activa.', 16, 1);
        END;

        SELECT
            @IdPlanCuentaGanancia = pc.IdPlanCuenta
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.CodigoCuenta = @CodigoCuentaGanancia
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1;

        SELECT
            @IdPlanCuentaPerdida = pc.IdPlanCuenta
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.CodigoCuenta = @CodigoCuentaPerdida
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1;

        IF @IdPlanCuentaGanancia IS NULL OR @IdPlanCuentaPerdida IS NULL
        BEGIN
            RAISERROR(N'Las cuentas CUENTAGANANCIA_AJ o CUENTAPERDIDA_AJ no existen o no aceptan movimiento en el plan de cuentas.', 16, 1);
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

        SELECT
            @IdMonedaUsd = m.IdMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.CodigoMoneda = 'USD'
          AND m.Estado = 1;

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        DECLARE @CuentasProceso TABLE
        (
            IdPlanCuenta INT NOT NULL PRIMARY KEY,
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            CodigoMonedaCuenta VARCHAR(3) NOT NULL,
            TipoCambioCuenta CHAR(1) NOT NULL
        );

        INSERT INTO @CuentasProceso
        (
            IdPlanCuenta,
            CodigoCuenta,
            NombreCuenta,
            CodigoMonedaCuenta,
            TipoCambioCuenta
        )
        SELECT
            pc.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            CASE
                WHEN pc.IdMoneda = 'USD' THEN 'USD'
                ELSE 'PEN'
            END,
            ISNULL(pc.TipoCambio, '')
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1
          AND pc.GeneraDiferenciaPorAnalisis = 1;

        SET @TotalCuentas = @@ROWCOUNT;

        IF EXISTS
        (
            SELECT 1
            FROM @CuentasProceso AS c
            WHERE c.CodigoMonedaCuenta = 'USD'
              AND c.TipoCambioCuenta NOT IN ('C', 'V')
        )
        BEGIN
            RAISERROR(N'Existen cuentas analiticas en USD sin tipo de cambio Compra/Venta configurado.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @CuentasProceso AS c
            WHERE c.CodigoMonedaCuenta = 'USD'
        )
        BEGIN
            SELECT
                @IdCuentaAdministradora = e.IdCuentaAdministradora
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa;

            IF @IdCuentaAdministradora IS NULL
            BEGIN
                RAISERROR(N'La empresa no tiene una cuenta administradora asociada para obtener el tipo de cambio.', 16, 1);
            END;

            SELECT
                @TipoCambioCompra = tc.Compra,
                @TipoCambioVenta = tc.Venta
            FROM dbo.CON_TipoCambio AS tc
            WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
              AND tc.Fecha = @FechaAsiento
              AND tc.IdMoneda = 'USD'
              AND tc.Estado = 1;

            IF @TipoCambioCompra <= 0 OR @TipoCambioVenta <= 0
            BEGIN
                RAISERROR(N'No existe tipo de cambio USD para la fecha de cierre del periodo seleccionado.', 16, 1);
            END;
        END;

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_AjusteCuentaProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo
        )
        BEGIN
            INSERT INTO @AsientosEliminar (IdAsiento)
            SELECT DISTINCT
                d.IdAsiento
            FROM dbo.CON_AjusteCuentaProcesoDetalle AS d
            INNER JOIN dbo.CON_AjusteCuentaProceso AS p
                ON p.IdAjusteCuentaProceso = d.IdAjusteCuentaProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo
              AND d.IdAsiento IS NOT NULL;

            DELETE d
            FROM dbo.CON_AjusteCuentaProcesoDetalle AS d
            INNER JOIN dbo.CON_AjusteCuentaProceso AS p
                ON p.IdAjusteCuentaProceso = d.IdAjusteCuentaProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo;

            DELETE p
            FROM dbo.CON_AjusteCuentaProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo;

            DELETE d
            FROM dbo.CON_AsientoDetalle AS d
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = d.IdAsiento;

            DELETE a
            FROM dbo.CON_Asiento AS a
            INNER JOIN @AsientosEliminar AS e
                ON e.IdAsiento = a.IdAsiento;

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_CorrelativoAsiento AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @Periodo
            )
            BEGIN
                DECLARE @UltimoNumeroRestante INT = 0;

                SELECT
                    @UltimoNumeroRestante = ISNULL(MAX(a.NumeroAsiento), 0)
                FROM dbo.CON_Asiento AS a
                WHERE a.IdEmpresa = @IdEmpresa
                  AND a.IdOrigen = @IdOrigen
                  AND a.Periodo = @Periodo;

                IF @UltimoNumeroRestante = 0
                BEGIN
                    DELETE dbo.CON_CorrelativoAsiento
                    WHERE IdEmpresa = @IdEmpresa
                      AND IdOrigen = @IdOrigen
                      AND Periodo = @Periodo;
                END
                ELSE
                BEGIN
                    UPDATE dbo.CON_CorrelativoAsiento
                    SET UltimoNumero = @UltimoNumeroRestante,
                        FechaActualizacion = SYSDATETIME(),
                        UsuarioRegistro = @UsuarioRegistro
                    WHERE IdEmpresa = @IdEmpresa
                      AND IdOrigen = @IdOrigen
                      AND Periodo = @Periodo;
                END;
            END;
        END;

        INSERT INTO dbo.CON_AjusteCuentaProceso
        (
            IdEmpresa,
            Periodo,
            IdOrigen,
            FechaAsiento,
            TotalCuentas,
            TotalAsientos,
            TotalDebe,
            TotalHaber,
            UsuarioRegistro
        )
        VALUES
        (
            @IdEmpresa,
            @Periodo,
            @IdOrigen,
            @FechaAsiento,
            @TotalCuentas,
            0,
            0,
            0,
            @UsuarioRegistro
        );

        SET @IdAjusteCuentaProceso = SCOPE_IDENTITY();

        DECLARE
            @IdPlanCuentaTrabajo INT,
            @CodigoCuentaTrabajo VARCHAR(20),
            @NombreCuentaTrabajo NVARCHAR(200),
            @CodigoMonedaCuenta VARCHAR(3),
            @TipoCambioCuenta CHAR(1),
            @TipoCambioAplicado DECIMAL(18,6),
            @IdMonedaAsiento INT,
            @IdAsientoTrabajo INT,
            @NumeroAsientoTrabajo INT,
            @TotalDebeCuenta DECIMAL(18,2),
            @TotalHaberCuenta DECIMAL(18,2),
            @TotalDebeCuentaSoles DECIMAL(18,2),
            @TotalHaberCuentaSoles DECIMAL(18,2),
            @TotalAnalisisCuenta INT,
            @GlosaAsiento NVARCHAR(500),
            @ObservacionDetalle NVARCHAR(300),
            @AplicoCuentaDestino BIT;

        DECLARE cursor_cuentas CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            c.IdPlanCuenta,
            c.CodigoCuenta,
            c.NombreCuenta,
            c.CodigoMonedaCuenta,
            c.TipoCambioCuenta
        FROM @CuentasProceso AS c
        ORDER BY
            c.CodigoCuenta ASC;

        OPEN cursor_cuentas;

        FETCH NEXT FROM cursor_cuentas
        INTO @IdPlanCuentaTrabajo, @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @CodigoMonedaCuenta, @TipoCambioCuenta;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @TipoCambioAplicado = CASE
                                          WHEN @CodigoMonedaCuenta = 'USD' AND @TipoCambioCuenta = 'C' THEN @TipoCambioCompra
                                          WHEN @CodigoMonedaCuenta = 'USD' AND @TipoCambioCuenta = 'V' THEN @TipoCambioVenta
                                          ELSE 1
                                      END;
            SET @IdAsientoTrabajo = NULL;
            SET @NumeroAsientoTrabajo = NULL;
            SET @IdMonedaAsiento = CASE WHEN @CodigoMonedaCuenta = 'USD' THEN @IdMonedaUsd ELSE @IdMonedaPen END;
            SET @TotalDebeCuenta = 0;
            SET @TotalHaberCuenta = 0;
            SET @TotalDebeCuentaSoles = 0;
            SET @TotalHaberCuentaSoles = 0;
            SET @TotalAnalisisCuenta = 0;
            SET @ObservacionDetalle = NULL;
            SET @AplicoCuentaDestino = 0;

            IF @CodigoMonedaCuenta = 'USD' AND @IdMonedaAsiento IS NULL
            BEGIN
                RAISERROR(N'La moneda USD no esta registrada como activa.', 16, 1);
            END;

            DECLARE @AnalisisCuenta TABLE
            (
                Item INT IDENTITY(1,1) NOT NULL,
                IdCliente INT NULL,
                IdProveedor INT NULL,
                NumeroDocumento VARCHAR(20) NULL,
                TipoDocumento NVARCHAR(150) NULL,
                Serie VARCHAR(10) NULL,
                ReferenciaLinea NVARCHAR(100) NULL,
                ResiduoMoneda DECIMAL(18,2) NOT NULL,
                ImporteMoneda DECIMAL(18,2) NOT NULL,
                ImporteSoles DECIMAL(18,2) NOT NULL
            );

            DELETE FROM @AnalisisCuenta;

            INSERT INTO @AnalisisCuenta
            (
                IdCliente,
                IdProveedor,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                ResiduoMoneda,
                ImporteMoneda,
                ImporteSoles
            )
            SELECT
                NULL,
                NULL,
                NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(d.Serie)), ''),
                NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''),
                ROUND(
                    SUM(
                        CASE
                            WHEN d.DH = 'D' THEN
                                CASE
                                    WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                    ELSE d.TotalImporteS
                                END
                            ELSE 0
                        END
                    ) -
                    SUM(
                        CASE
                            WHEN d.DH = 'H' THEN
                                CASE
                                    WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                    ELSE d.TotalImporteS
                                END
                            ELSE 0
                        END
                    ),
                    2
                ) AS ResiduoMoneda,
                ABS(
                    ROUND(
                        SUM(
                            CASE
                                WHEN d.DH = 'D' THEN
                                    CASE
                                        WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                        ELSE d.TotalImporteS
                                    END
                                ELSE 0
                            END
                        ) -
                        SUM(
                            CASE
                                WHEN d.DH = 'H' THEN
                                    CASE
                                        WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                        ELSE d.TotalImporteS
                                    END
                                ELSE 0
                            END
                        ),
                        2
                    )
                ),
                CASE
                    WHEN @CodigoMonedaCuenta = 'USD' THEN ROUND(
                        ABS(
                            ROUND(
                                SUM(
                                    CASE
                                        WHEN d.DH = 'D' THEN d.TotalImporteD
                                        ELSE 0
                                    END
                                ) -
                                SUM(
                                    CASE
                                        WHEN d.DH = 'H' THEN d.TotalImporteD
                                        ELSE 0
                                    END
                                ),
                                2
                            )
                        ) * @TipoCambioAplicado,
                        2
                    )
                    ELSE ABS(
                        ROUND(
                            SUM(
                                CASE
                                    WHEN d.DH = 'D' THEN d.TotalImporteS
                                    ELSE 0
                                END
                            ) -
                            SUM(
                                CASE
                                    WHEN d.DH = 'H' THEN d.TotalImporteS
                                    ELSE 0
                                END
                            ),
                            2
                        )
                    )
                END
            FROM dbo.CON_AsientoDetalle AS d
            INNER JOIN dbo.CON_Asiento AS a
                ON a.IdAsiento = d.IdAsiento
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.Periodo <= @Periodo
              AND LEFT(a.Periodo, 4) = LEFT(@Periodo, 4)
              AND a.IdOrigen <> @IdOrigen
              AND d.IdPlanCuenta = @IdPlanCuentaTrabajo
              AND NOT EXISTS
              (
                  SELECT 1
                  FROM dbo.CON_ConfiguracionContabilizacion AS cfg
                  WHERE cfg.IdEmpresa = @IdEmpresa
                    AND cfg.IdOrigen = a.IdOrigen
                    AND cfg.Activo = 1
                    AND cfg.ModuloOperacion IN ('DIF', 'AJU', 'APR', 'CIE')
              )
              AND NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), '') IS NOT NULL
            GROUP BY
                NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(d.Serie)), ''),
                NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), '')
            HAVING ABS(
                       ROUND(
                           SUM(
                               CASE
                                   WHEN d.DH = 'D' THEN
                                       CASE
                                           WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                           ELSE d.TotalImporteS
                                       END
                                   ELSE 0
                               END
                           ) -
                           SUM(
                               CASE
                                   WHEN d.DH = 'H' THEN
                                       CASE
                                           WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                           ELSE d.TotalImporteS
                                       END
                                   ELSE 0
                               END
                           ),
                           2
                       )
                   ) > 0.004
               AND ABS(
                       ROUND(
                           SUM(
                               CASE
                                   WHEN d.DH = 'D' THEN
                                       CASE
                                           WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                           ELSE d.TotalImporteS
                                       END
                                   ELSE 0
                               END
                           ) -
                           SUM(
                               CASE
                                   WHEN d.DH = 'H' THEN
                                       CASE
                                           WHEN @CodigoMonedaCuenta = 'USD' THEN d.TotalImporteD
                                           ELSE d.TotalImporteS
                                       END
                                   ELSE 0
                               END
                           ),
                           2
                       )
                   ) < 1;

            SELECT
                @TotalAnalisisCuenta = COUNT(*)
            FROM @AnalisisCuenta;

            DECLARE @LineasCuenta TABLE
            (
                Item INT IDENTITY(1,1) NOT NULL,
                IdPlanCuenta INT NOT NULL,
                GlosaDetalle NVARCHAR(300) NOT NULL,
                NumeroDocumento VARCHAR(20) NULL,
                TipoDocumento NVARCHAR(150) NULL,
                Serie VARCHAR(10) NULL,
                ReferenciaLinea NVARCHAR(100) NULL,
                IdCliente INT NULL,
                IdProveedor INT NULL,
                TipoCambioLinea DECIMAL(18,6) NOT NULL,
                Debe DECIMAL(18,2) NOT NULL,
                Haber DECIMAL(18,2) NOT NULL,
                TotalImporteS DECIMAL(18,2) NOT NULL,
                TotalImporteD DECIMAL(18,2) NOT NULL
            );

            DELETE FROM @LineasCuenta;

            IF @TotalAnalisisCuenta > 0
            BEGIN
                INSERT INTO @LineasCuenta
                (
                    IdPlanCuenta,
                    GlosaDetalle,
                    NumeroDocumento,
                    TipoDocumento,
                    Serie,
                    ReferenciaLinea,
                    IdCliente,
                    IdProveedor,
                    TipoCambioLinea,
                    Debe,
                    Haber,
                    TotalImporteS,
                    TotalImporteD
                )
                SELECT
                    @IdPlanCuentaTrabajo,
                    LEFT(
                        CONCAT(
                            N'AJUSTE ',
                            @CodigoCuentaTrabajo,
                            CASE
                                WHEN a.NumeroDocumento IS NOT NULL THEN CONCAT(N' AUX ', a.NumeroDocumento)
                                ELSE N''
                            END,
                            CASE
                                WHEN a.TipoDocumento IS NOT NULL OR a.Serie IS NOT NULL OR a.ReferenciaLinea IS NOT NULL
                                    THEN CONCAT(
                                        N' / ',
                                        COALESCE(a.TipoDocumento, N''),
                                        CASE WHEN a.TipoDocumento IS NOT NULL AND a.Serie IS NOT NULL THEN N' ' ELSE N'' END,
                                        COALESCE(a.Serie, ''),
                                        CASE WHEN a.Serie IS NOT NULL AND a.ReferenciaLinea IS NOT NULL THEN N'-' ELSE N'' END,
                                        COALESCE(a.ReferenciaLinea, N'')
                                    )
                                ELSE N''
                            END
                        ),
                        300
                    ),
                    a.NumeroDocumento,
                    a.TipoDocumento,
                    a.Serie,
                    a.ReferenciaLinea,
                    NULL,
                    NULL,
                    @TipoCambioAplicado,
                    CASE
                        WHEN a.ResiduoMoneda < 0 THEN
                            CASE WHEN @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE a.ImporteSoles END
                        ELSE 0
                    END,
                    CASE
                        WHEN a.ResiduoMoneda > 0 THEN
                            CASE WHEN @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE a.ImporteSoles END
                        ELSE 0
                    END,
                    a.ImporteSoles,
                    CASE WHEN @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE 0 END
                FROM @AnalisisCuenta AS a
                ORDER BY
                    a.Item ASC;

                DECLARE @AcumuladoPerdidaSoles DECIMAL(18,2) = 0
                DECLARE @AcumuladoPerdidaDolares DECIMAL(18,2) = 0
                DECLARE @AcumuladoGananciaSoles DECIMAL(18,2) = 0
                DECLARE @AcumuladoGananciaDolares DECIMAL(18,2) = 0
                DECLARE @AcumuladoPerdidaAsiento DECIMAL(18,2) = 0
                DECLARE @AcumuladoGananciaAsiento DECIMAL(18,2) = 0

                SELECT
                    @AcumuladoPerdidaSoles = ISNULL(SUM(CASE WHEN a.ResiduoMoneda > 0 THEN a.ImporteSoles ELSE 0 END), 0),
                    @AcumuladoPerdidaDolares = ISNULL(SUM(CASE WHEN a.ResiduoMoneda > 0 AND @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE 0 END), 0),
                    @AcumuladoGananciaSoles = ISNULL(SUM(CASE WHEN a.ResiduoMoneda < 0 THEN a.ImporteSoles ELSE 0 END), 0),
                    @AcumuladoGananciaDolares = ISNULL(SUM(CASE WHEN a.ResiduoMoneda < 0 AND @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE 0 END), 0),
                    @AcumuladoPerdidaAsiento = ISNULL(SUM(CASE WHEN a.ResiduoMoneda > 0 THEN CASE WHEN @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE a.ImporteSoles END ELSE 0 END), 0),
                    @AcumuladoGananciaAsiento = ISNULL(SUM(CASE WHEN a.ResiduoMoneda < 0 THEN CASE WHEN @CodigoMonedaCuenta = 'USD' THEN a.ImporteMoneda ELSE a.ImporteSoles END ELSE 0 END), 0)
                FROM @AnalisisCuenta AS a;

                IF @AcumuladoPerdidaAsiento > 0
                BEGIN
                    INSERT INTO @LineasCuenta
                    (
                        IdPlanCuenta,
                        GlosaDetalle,
                        NumeroDocumento,
                        TipoDocumento,
                        Serie,
                        ReferenciaLinea,
                        IdCliente,
                        IdProveedor,
                        TipoCambioLinea,
                        Debe,
                        Haber,
                        TotalImporteS,
                        TotalImporteD
                    )
                    VALUES
                    (
                        @IdPlanCuentaPerdida,
                        N'PERDIDA POR AJUSTE DE CUENTAS',
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        @TipoCambioAplicado,
                        @AcumuladoPerdidaAsiento,
                        0,
                        @AcumuladoPerdidaSoles,
                        @AcumuladoPerdidaDolares
                    );
                END;

                IF @AcumuladoGananciaAsiento > 0
                BEGIN
                    INSERT INTO @LineasCuenta
                    (
                        IdPlanCuenta,
                        GlosaDetalle,
                        NumeroDocumento,
                        TipoDocumento,
                        Serie,
                        ReferenciaLinea,
                        IdCliente,
                        IdProveedor,
                        TipoCambioLinea,
                        Debe,
                        Haber,
                        TotalImporteS,
                        TotalImporteD
                    )
                    VALUES
                    (
                        @IdPlanCuentaGanancia,
                        N'GANANCIA POR AJUSTE DE CUENTAS',
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        NULL,
                        @TipoCambioAplicado,
                        0,
                        @AcumuladoGananciaAsiento,
                        @AcumuladoGananciaSoles,
                        @AcumuladoGananciaDolares
                    );
                END;
            END;

            DECLARE @CuentaDestinoDetalle TABLE
            (
                IdPlanCuentaOrigen INT NOT NULL,
                Orden SMALLINT NOT NULL,
                IdPlanCuentaDestinoCargo INT NOT NULL,
                IdPlanCuentaDestinoAbono INT NOT NULL,
                Porcentaje DECIMAL(7,4) NOT NULL,
                EsUltimo BIT NOT NULL
            );

            DELETE FROM @CuentaDestinoDetalle;

            INSERT INTO @CuentaDestinoDetalle
            (
                IdPlanCuentaOrigen,
                Orden,
                IdPlanCuentaDestinoCargo,
                IdPlanCuentaDestinoAbono,
                Porcentaje,
                EsUltimo
            )
            SELECT
                r.IdPlanCuentaOrigen,
                d.Orden,
                d.IdPlanCuentaDestinoCargo,
                d.IdPlanCuentaDestinoAbono,
                d.Porcentaje,
                CASE
                    WHEN ROW_NUMBER() OVER (PARTITION BY r.IdPlanCuentaOrigen ORDER BY d.Orden DESC) = 1 THEN 1
                    ELSE 0
                END
            FROM dbo.CON_CuentaDestinoRegla AS r
            INNER JOIN dbo.CON_CuentaDestinoReglaDetalle AS d
                ON d.IdCuentaDestinoRegla = r.IdCuentaDestinoRegla
               AND d.Activo = 1
            INNER JOIN
            (
                SELECT DISTINCT
                    l.IdPlanCuenta
                FROM @LineasCuenta AS l
            ) AS base
                ON base.IdPlanCuenta = r.IdPlanCuentaOrigen
            WHERE r.IdEmpresa = @IdEmpresa
              AND r.Activo = 1;

            IF EXISTS
            (
                SELECT 1
                FROM @CuentaDestinoDetalle AS d
                LEFT JOIN dbo.CON_PlanCuenta AS cargo
                    ON cargo.IdPlanCuenta = d.IdPlanCuentaDestinoCargo
                   AND cargo.IdEmpresa = @IdEmpresa
                   AND cargo.Estado = 1
                   AND cargo.AceptaMovimiento = 1
                LEFT JOIN dbo.CON_PlanCuenta AS abono
                    ON abono.IdPlanCuenta = d.IdPlanCuentaDestinoAbono
                   AND abono.IdEmpresa = @IdEmpresa
                   AND abono.Estado = 1
                   AND abono.AceptaMovimiento = 1
                WHERE cargo.IdPlanCuenta IS NULL
                   OR abono.IdPlanCuenta IS NULL
            )
            BEGIN
                RAISERROR(N'Existe una configuracion activa de cuentas destino con cuentas cargo o abono invalidas para la empresa.', 16, 1);
            END;

            IF EXISTS
            (
                SELECT 1
                FROM @CuentaDestinoDetalle
            )
            BEGIN
                DECLARE @IdPlanCuentaOrigenDestino INT
                DECLARE @ItemLineaDestino INT
                DECLARE @GlosaBaseDestino NVARCHAR(300)
                DECLARE @NumeroDocumentoDestino VARCHAR(20)
                DECLARE @TipoDocumentoDestino NVARCHAR(150)
                DECLARE @SerieDestino VARCHAR(10)
                DECLARE @ReferenciaLineaDestino NVARCHAR(100)
                DECLARE @IdClienteDestino INT
                DECLARE @IdProveedorDestino INT
                DECLARE @TipoCambioLineaDestino DECIMAL(18,6)
                DECLARE @DebeOrigenDestino DECIMAL(18,2)
                DECLARE @HaberOrigenDestino DECIMAL(18,2)
                DECLARE @TotalImporteSOrigenDestino DECIMAL(18,2)
                DECLARE @TotalImporteDOrigenDestino DECIMAL(18,2)
                DECLARE @IdCuentaCargoDestino INT
                DECLARE @IdCuentaAbonoDestino INT
                DECLARE @PorcentajeDestino DECIMAL(7,4)
                DECLARE @EsUltimoDestino BIT
                DECLARE @ImporteBaseDestinoAsiento DECIMAL(18,2)
                DECLARE @ImporteBaseDestinoS DECIMAL(18,2)
                DECLARE @ImporteBaseDestinoD DECIMAL(18,2)
                DECLARE @ImporteDistribuidoDestinoAsiento DECIMAL(18,2)
                DECLARE @ImporteDistribuidoDestinoS DECIMAL(18,2)
                DECLARE @ImporteDistribuidoDestinoD DECIMAL(18,2)
                DECLARE @ImporteTramoDestinoAsiento DECIMAL(18,2)
                DECLARE @ImporteTramoDestinoS DECIMAL(18,2)
                DECLARE @ImporteTramoDestinoD DECIMAL(18,2)

                DECLARE cursor_linea_destino CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    l.IdPlanCuenta,
                    l.Item,
                    l.GlosaDetalle,
                    l.NumeroDocumento,
                    l.TipoDocumento,
                    l.Serie,
                    l.ReferenciaLinea,
                    l.IdCliente,
                    l.IdProveedor,
                    l.TipoCambioLinea,
                    l.Debe,
                    l.Haber,
                    l.TotalImporteS,
                    l.TotalImporteD
                FROM @LineasCuenta AS l
                WHERE (l.Debe > 0 OR l.Haber > 0)
                  AND EXISTS
                  (
                      SELECT 1
                      FROM @CuentaDestinoDetalle AS r
                      WHERE r.IdPlanCuentaOrigen = l.IdPlanCuenta
                  )
                ORDER BY
                    l.Item ASC;

                OPEN cursor_linea_destino;

                FETCH NEXT FROM cursor_linea_destino
                INTO @IdPlanCuentaOrigenDestino, @ItemLineaDestino, @GlosaBaseDestino, @NumeroDocumentoDestino, @TipoDocumentoDestino, @SerieDestino,
                     @ReferenciaLineaDestino, @IdClienteDestino, @IdProveedorDestino, @TipoCambioLineaDestino,
                     @DebeOrigenDestino, @HaberOrigenDestino, @TotalImporteSOrigenDestino, @TotalImporteDOrigenDestino;

                WHILE @@FETCH_STATUS = 0
                BEGIN
                    SET @ImporteBaseDestinoAsiento = CASE
                                                         WHEN @DebeOrigenDestino > 0 THEN @DebeOrigenDestino
                                                         ELSE @HaberOrigenDestino
                                                     END;
                    SET @ImporteBaseDestinoS = @TotalImporteSOrigenDestino;
                    SET @ImporteBaseDestinoD = @TotalImporteDOrigenDestino;
                    SET @ImporteDistribuidoDestinoAsiento = 0;
                    SET @ImporteDistribuidoDestinoS = 0;
                    SET @ImporteDistribuidoDestinoD = 0;

                    DECLARE cursor_tramo_destino CURSOR LOCAL FAST_FORWARD FOR
                    SELECT
                        r.IdPlanCuentaDestinoCargo,
                        r.IdPlanCuentaDestinoAbono,
                        r.Porcentaje,
                        r.EsUltimo
                    FROM @CuentaDestinoDetalle AS r
                    WHERE r.IdPlanCuentaOrigen = @IdPlanCuentaOrigenDestino
                    ORDER BY
                        r.Orden ASC;

                    OPEN cursor_tramo_destino;

                    FETCH NEXT FROM cursor_tramo_destino
                    INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;

                    WHILE @@FETCH_STATUS = 0
                    BEGIN
                        SET @ImporteTramoDestinoAsiento = CASE
                                                              WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDestinoAsiento - @ImporteDistribuidoDestinoAsiento
                                                              ELSE ROUND(@ImporteBaseDestinoAsiento * (@PorcentajeDestino / 100.0), 2)
                                                          END;
                        SET @ImporteTramoDestinoS = CASE
                                                        WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDestinoS - @ImporteDistribuidoDestinoS
                                                        ELSE ROUND(@ImporteBaseDestinoS * (@PorcentajeDestino / 100.0), 2)
                                                    END;
                        SET @ImporteTramoDestinoD = CASE
                                                        WHEN @ImporteBaseDestinoD = 0 THEN 0
                                                        WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDestinoD - @ImporteDistribuidoDestinoD
                                                        ELSE ROUND(@ImporteBaseDestinoD * (@PorcentajeDestino / 100.0), 2)
                                                    END;

                        IF @ImporteTramoDestinoAsiento <> 0
                        BEGIN
                            INSERT INTO @LineasCuenta
                            (
                                IdPlanCuenta,
                                GlosaDetalle,
                                NumeroDocumento,
                                TipoDocumento,
                                Serie,
                                ReferenciaLinea,
                                IdCliente,
                                IdProveedor,
                                TipoCambioLinea,
                                Debe,
                                Haber,
                                TotalImporteS,
                                TotalImporteD
                            )
                            VALUES
                            (
                                @IdCuentaCargoDestino,
                                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'AJUSTE DE CUENTAS'), N' / Destino'), 300),
                                @NumeroDocumentoDestino,
                                @TipoDocumentoDestino,
                                @SerieDestino,
                                @ReferenciaLineaDestino,
                                NULL,
                                NULL,
                                @TipoCambioLineaDestino,
                                @ImporteTramoDestinoAsiento,
                                0,
                                @ImporteTramoDestinoS,
                                @ImporteTramoDestinoD
                            );

                            INSERT INTO @LineasCuenta
                            (
                                IdPlanCuenta,
                                GlosaDetalle,
                                NumeroDocumento,
                                TipoDocumento,
                                Serie,
                                ReferenciaLinea,
                                IdCliente,
                                IdProveedor,
                                TipoCambioLinea,
                                Debe,
                                Haber,
                                TotalImporteS,
                                TotalImporteD
                            )
                            VALUES
                            (
                                @IdCuentaAbonoDestino,
                                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'AJUSTE DE CUENTAS'), N' / Contrapartida'), 300),
                                @NumeroDocumentoDestino,
                                @TipoDocumentoDestino,
                                @SerieDestino,
                                @ReferenciaLineaDestino,
                                NULL,
                                NULL,
                                @TipoCambioLineaDestino,
                                0,
                                @ImporteTramoDestinoAsiento,
                                @ImporteTramoDestinoS,
                                @ImporteTramoDestinoD
                            );

                            SET @AplicoCuentaDestino = 1;
                        END;

                        SET @ImporteDistribuidoDestinoAsiento = @ImporteDistribuidoDestinoAsiento + @ImporteTramoDestinoAsiento;
                        SET @ImporteDistribuidoDestinoS = @ImporteDistribuidoDestinoS + @ImporteTramoDestinoS;
                        SET @ImporteDistribuidoDestinoD = @ImporteDistribuidoDestinoD + @ImporteTramoDestinoD;

                        FETCH NEXT FROM cursor_tramo_destino
                        INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
                    END;

                    CLOSE cursor_tramo_destino;
                    DEALLOCATE cursor_tramo_destino;

                    FETCH NEXT FROM cursor_linea_destino
                    INTO @IdPlanCuentaOrigenDestino, @ItemLineaDestino, @GlosaBaseDestino, @NumeroDocumentoDestino, @TipoDocumentoDestino, @SerieDestino,
                         @ReferenciaLineaDestino, @IdClienteDestino, @IdProveedorDestino, @TipoCambioLineaDestino,
                         @DebeOrigenDestino, @HaberOrigenDestino, @TotalImporteSOrigenDestino, @TotalImporteDOrigenDestino;
                END;

                CLOSE cursor_linea_destino;
                DEALLOCATE cursor_linea_destino;
            END;

            SELECT
                @TotalDebeCuenta = ISNULL(SUM(l.Debe), 0),
                @TotalHaberCuenta = ISNULL(SUM(l.Haber), 0),
                @TotalDebeCuentaSoles = ISNULL(SUM(CASE WHEN l.Debe > 0 THEN l.TotalImporteS ELSE 0 END), 0),
                @TotalHaberCuentaSoles = ISNULL(SUM(CASE WHEN l.Haber > 0 THEN l.TotalImporteS ELSE 0 END), 0)
            FROM @LineasCuenta AS l;

            IF EXISTS
            (
                SELECT 1
                FROM @LineasCuenta AS l
            )
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
                        @NumeroAsientoTrabajo = c.UltimoNumero
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

                    SET @NumeroAsientoTrabajo = 1;
                END;

                SET @GlosaAsiento = CONCAT(N'AJUSTE DE CUENTAS ', @CodigoCuentaTrabajo, N' - ', @NombreCuentaTrabajo);

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
                    @Mes,
                    @Periodo,
                    @NumeroAsientoTrabajo,
                    @FechaAsiento,
                    @FechaAsiento,
                    @GlosaAsiento,
                    @IdMonedaAsiento,
                    @TipoCambioAplicado,
                    @TotalDebeCuenta,
                    @TotalHaberCuenta,
                    N'PROVISIONADO',
                    CONCAT(N'AJU-', @Periodo, N'-', @CodigoCuentaTrabajo),
                    CASE
                        WHEN @AplicoCuentaDestino = 1 THEN N'Generado por analisis con expansion de cuentas destino.'
                        ELSE N'Generado por analisis.'
                    END,
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
                    TipoDocumento,
                    NumeroDocumento,
                    Serie,
                    TipoCambioLinea,
                    IdCliente,
                    IdProveedor,
                    Debe,
                    Haber,
                    TotalImporteS,
                    TotalImporteD,
                    ReferenciaLinea,
                    UsuarioRegistro
                )
                SELECT
                    @IdAsientoTrabajo,
                    l.Item,
                    l.IdPlanCuenta,
                    CASE WHEN l.Debe > 0 THEN 'D' ELSE 'H' END,
                    l.GlosaDetalle,
                    l.TipoDocumento,
                    l.NumeroDocumento,
                    l.Serie,
                    l.TipoCambioLinea,
                    l.IdCliente,
                    l.IdProveedor,
                    l.Debe,
                    l.Haber,
                    l.TotalImporteS,
                    l.TotalImporteD,
                    l.ReferenciaLinea,
                    @UsuarioRegistro
                FROM @LineasCuenta AS l
                ORDER BY
                    l.Item ASC;

                SET @TotalAsientos += 1;
                SET @TotalDebeProceso += @TotalDebeCuentaSoles;
                SET @TotalHaberProceso += @TotalHaberCuentaSoles;
                SET @ObservacionDetalle = CASE
                                              WHEN @AplicoCuentaDestino = 1 THEN N'Asiento generado por analisis documental con expansion de cuentas destino.'
                                              ELSE N'Asiento generado por analisis documental.'
                                          END;

                INSERT INTO dbo.CON_AjusteCuentaProcesoDetalle
                (
                    IdAjusteCuentaProceso,
                    IdPlanCuenta,
                    CodigoMoneda,
                    TipoCambioAplicado,
                    TotalAnalisis,
                    IdAsiento,
                    NumeroAsiento,
                    TotalDebe,
                    TotalHaber,
                    Estado,
                    Observacion,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdAjusteCuentaProceso,
                    @IdPlanCuentaTrabajo,
                    @CodigoMonedaCuenta,
                    @TipoCambioAplicado,
                    @TotalAnalisisCuenta,
                    @IdAsientoTrabajo,
                    @NumeroAsientoTrabajo,
                    @TotalDebeCuenta,
                    @TotalHaberCuenta,
                    N'GENERADO',
                    @ObservacionDetalle,
                    @UsuarioRegistro
                );
            END
            ELSE
            BEGIN
                INSERT INTO dbo.CON_AjusteCuentaProcesoDetalle
                (
                    IdAjusteCuentaProceso,
                    IdPlanCuenta,
                    CodigoMoneda,
                    TipoCambioAplicado,
                    TotalAnalisis,
                    IdAsiento,
                    NumeroAsiento,
                    TotalDebe,
                    TotalHaber,
                    Estado,
                    Observacion,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdAjusteCuentaProceso,
                    @IdPlanCuentaTrabajo,
                    @CodigoMonedaCuenta,
                    @TipoCambioAplicado,
                    0,
                    NULL,
                    NULL,
                    0,
                    0,
                    N'SIN_AJUSTE',
                    N'La cuenta no genero residuales por analisis menores a una unidad para el periodo seleccionado.',
                    @UsuarioRegistro
                );
            END;

            FETCH NEXT FROM cursor_cuentas
            INTO @IdPlanCuentaTrabajo, @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @CodigoMonedaCuenta, @TipoCambioCuenta;
        END;

        CLOSE cursor_cuentas;
        DEALLOCATE cursor_cuentas;

        UPDATE dbo.CON_AjusteCuentaProceso
        SET TotalCuentas = @TotalCuentas,
            TotalAsientos = @TotalAsientos,
            TotalDebe = @TotalDebeProceso,
            TotalHaber = @TotalHaberProceso,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdAjusteCuentaProceso = @IdAjusteCuentaProceso;

        COMMIT;
        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        EXEC dbo.usp_CON_ObtenerAjusteCuentaProceso
            @IdEmpresa = @IdEmpresa,
            @Periodo = @Periodo;

    END TRY

    BEGIN CATCH

        IF CURSOR_STATUS('local', 'cursor_cuentas') >= -1
        BEGIN
            CLOSE cursor_cuentas;
            DEALLOCATE cursor_cuentas;
        END;

        IF CURSOR_STATUS('local', 'cursor_linea_destino') >= -1
        BEGIN
            CLOSE cursor_linea_destino;
            DEALLOCATE cursor_linea_destino;
        END;

        IF CURSOR_STATUS('local', 'cursor_tramo_destino') >= -1
        BEGIN
            CLOSE cursor_tramo_destino;
            DEALLOCATE cursor_tramo_destino;
        END;

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
