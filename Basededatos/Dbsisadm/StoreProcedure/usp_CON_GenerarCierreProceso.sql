-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Genera o regenera los asientos de cierre anual por cuenta usando ColBalance R/I.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Crea un proceso anual CIE con opcion de regenerar cierres 14 y/o 15, tomando cuentas por ColBalance y usando el TC editable de 31/12.
-- Firma: FRANCO LARA - 03/07/2026 | Usa DH como sentido explicito para el calculo de saldos base y lo persiste en las dos lineas de cada asiento de cierre.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarCierreProceso
    @IdEmpresa INT,
    @Anio SMALLINT,
    @TipoCambioCompra DECIMAL(18,6),
    @TipoCambioVenta DECIMAL(18,6),
    @ProcesarGananciasPerdidas BIT,
    @ProcesarInventarios BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @FechaAsiento DATE = DATEFROMPARTS(@Anio, 12, 31)
        DECLARE @PeriodoBaseInicio CHAR(6) = CONCAT(@Anio, '00')
        DECLARE @PeriodoBaseFin CHAR(6) = CONCAT(@Anio, '13')
        DECLARE @PeriodoGanancias CHAR(6) = CONCAT(@Anio, '14')
        DECLARE @PeriodoInventarios CHAR(6) = CONCAT(@Anio, '15')
        DECLARE @IdOrigen INT
        DECLARE @UsaTipoCambioSbs BIT = 0
        DECLARE @IdCuentaAdministradora INT
        DECLARE @CodigoCuentaGanancia VARCHAR(20)
        DECLARE @CodigoCuentaPerdida VARCHAR(20)
        DECLARE @IdPlanCuentaGanancia INT
        DECLARE @IdPlanCuentaPerdida INT
        DECLARE @CodigoMonedaGanancia VARCHAR(3)
        DECLARE @CodigoMonedaPerdida VARCHAR(3)
        DECLARE @TipoCambioCuentaGanancia CHAR(1)
        DECLARE @TipoCambioCuentaPerdida CHAR(1)
        DECLARE @NombreCuentaGanancia NVARCHAR(200)
        DECLARE @NombreCuentaPerdida NVARCHAR(200)
        DECLARE @IdMonedaPen INT
        DECLARE @IdCierreProceso INT
        DECLARE @TotalCuentas INT = 0
        DECLARE @TotalAsientos INT = 0
        DECLARE @TotalDebeProceso DECIMAL(18,2) = 0
        DECLARE @TotalHaberProceso DECIMAL(18,2) = 0

        IF @Anio < 2000 OR @Anio > 9999
        BEGIN
            RAISERROR(N'El ejercicio indicado es invalido.', 16, 1);
        END;

        IF ISNULL(@TipoCambioCompra, 0) <= 0 OR ISNULL(@TipoCambioVenta, 0) <= 0
        BEGIN
            RAISERROR(N'Ingrese un tipo de cambio de cierre mayor a cero.', 16, 1);
        END;

        IF ISNULL(@ProcesarGananciasPerdidas, 0) = 0
           AND ISNULL(@ProcesarInventarios, 0) = 0
        BEGIN
            RAISERROR(N'Seleccione al menos un tipo de cierre a procesar.', 16, 1);
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
                                    WHEN UPPER(LTRIM(RTRIM(pe.ValorParametro))) = 'S'
                                        THEN 1
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

        SELECT
            @CodigoCuentaGanancia = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAGANANCIA'
          AND pe.Activo = 1;

        SELECT
            @CodigoCuentaPerdida = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAPERDIDA'
          AND pe.Activo = 1;

        IF @CodigoCuentaGanancia IS NULL OR @CodigoCuentaPerdida IS NULL
        BEGIN
            RAISERROR(N'Configure las cuentas CUENTAGANANCIA y CUENTAPERDIDA para la empresa activa.', 16, 1);
        END;

        SELECT
            @IdPlanCuentaGanancia = pc.IdPlanCuenta,
            @CodigoMonedaGanancia = pc.IdMoneda,
            @TipoCambioCuentaGanancia = ISNULL(pc.TipoCambio, ''),
            @NombreCuentaGanancia = pc.NombreCuenta
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.CodigoCuenta = @CodigoCuentaGanancia
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1;

        SELECT
            @IdPlanCuentaPerdida = pc.IdPlanCuenta,
            @CodigoMonedaPerdida = pc.IdMoneda,
            @TipoCambioCuentaPerdida = ISNULL(pc.TipoCambio, ''),
            @NombreCuentaPerdida = pc.NombreCuenta
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.CodigoCuenta = @CodigoCuentaPerdida
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1;

        IF @IdPlanCuentaGanancia IS NULL OR @IdPlanCuentaPerdida IS NULL
        BEGIN
            RAISERROR(N'Las cuentas CUENTAGANANCIA o CUENTAPERDIDA no existen o no aceptan movimiento en el plan de cuentas.', 16, 1);
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

        DECLARE @CuentasProceso TABLE
        (
            TipoCierre CHAR(2) NOT NULL,
            DescripcionCierre NVARCHAR(100) NOT NULL,
            PeriodoTrabajo CHAR(6) NOT NULL,
            MesTrabajo TINYINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            CodigoMoneda VARCHAR(3) NOT NULL,
            TipoCambioCuenta CHAR(1) NOT NULL
        );

        IF @ProcesarGananciasPerdidas = 1
        BEGIN
            INSERT INTO @CuentasProceso
            (
                TipoCierre,
                DescripcionCierre,
                PeriodoTrabajo,
                MesTrabajo,
                IdPlanCuenta,
                CodigoCuenta,
                NombreCuenta,
                CodigoMoneda,
                TipoCambioCuenta
            )
            SELECT
                '14',
                N'Cierre de Ganancias y Perdidas',
                @PeriodoGanancias,
                14,
                pc.IdPlanCuenta,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                pc.IdMoneda,
                ISNULL(pc.TipoCambio, '')
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND pc.AceptaMovimiento = 1
              AND pc.ColBalance = 'R';
        END;

        IF @ProcesarInventarios = 1
        BEGIN
            INSERT INTO @CuentasProceso
            (
                TipoCierre,
                DescripcionCierre,
                PeriodoTrabajo,
                MesTrabajo,
                IdPlanCuenta,
                CodigoCuenta,
                NombreCuenta,
                CodigoMoneda,
                TipoCambioCuenta
            )
            SELECT
                '15',
                N'Cierre de Inventarios',
                @PeriodoInventarios,
                15,
                pc.IdPlanCuenta,
                pc.CodigoCuenta,
                pc.NombreCuenta,
                pc.IdMoneda,
                ISNULL(pc.TipoCambio, '')
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND pc.AceptaMovimiento = 1
              AND pc.ColBalance = 'I';
        END;

        SET @TotalCuentas = (SELECT COUNT(*) FROM @CuentasProceso);

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

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
                d.IdAsiento
            FROM dbo.CON_CierreProcesoDetalle AS d
            INNER JOIN dbo.CON_CierreProceso AS p
                ON p.IdCierreProceso = d.IdCierreProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Anio = @Anio
              AND d.IdAsiento IS NOT NULL;

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

            DELETE correlativo
            FROM dbo.CON_CorrelativoAsiento AS correlativo
            WHERE correlativo.IdEmpresa = @IdEmpresa
              AND correlativo.IdOrigen = @IdOrigen
              AND correlativo.Periodo IN (@PeriodoGanancias, @PeriodoInventarios)
              AND NOT EXISTS
              (
                  SELECT 1
                  FROM dbo.CON_Asiento AS a
                  WHERE a.IdEmpresa = @IdEmpresa
                    AND a.IdOrigen = @IdOrigen
                    AND a.Periodo = correlativo.Periodo
              );
        END;

        INSERT INTO dbo.CON_CierreProceso
        (
            IdEmpresa,
            Anio,
            IdOrigen,
            FechaAsiento,
            UsaTipoCambioSbs,
            TipoCambioCompra,
            TipoCambioVenta,
            ProcesaGananciasPerdidas,
            ProcesaInventarios,
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
            @IdOrigen,
            @FechaAsiento,
            @UsaTipoCambioSbs,
            @TipoCambioCompra,
            @TipoCambioVenta,
            @ProcesarGananciasPerdidas,
            @ProcesarInventarios,
            @TotalCuentas,
            0,
            0,
            0,
            @UsuarioRegistro
        );

        SET @IdCierreProceso = SCOPE_IDENTITY();

        DECLARE
            @TipoCierreTrabajo CHAR(2),
            @DescripcionCierreTrabajo NVARCHAR(100),
            @PeriodoTrabajo CHAR(6),
            @MesTrabajo TINYINT,
            @IdPlanCuentaTrabajo INT,
            @CodigoCuentaTrabajo VARCHAR(20),
            @NombreCuentaTrabajo NVARCHAR(200),
            @CodigoMonedaTrabajo VARCHAR(3),
            @TipoCambioCuentaTrabajo CHAR(1),
            @SaldoSoles DECIMAL(18,2),
            @SaldoDolares DECIMAL(18,2),
            @ImporteAbs DECIMAL(18,2),
            @IdPlanCuentaContra INT,
            @CodigoCuentaContra VARCHAR(20),
            @NombreCuentaContra NVARCHAR(200),
            @CodigoMonedaContra VARCHAR(3),
            @TipoCambioCuentaContra CHAR(1),
            @TipoCambioAplicado DECIMAL(18,6),
            @TipoCambioContraAplicado DECIMAL(18,6),
            @TotalImporteDLinea1 DECIMAL(18,2),
            @TotalImporteDLinea2 DECIMAL(18,2),
            @NumeroAsientoTrabajo INT,
            @IdAsientoTrabajo INT,
            @GlosaAsiento NVARCHAR(500),
            @DebeLinea1 DECIMAL(18,2),
            @HaberLinea1 DECIMAL(18,2),
            @DebeLinea2 DECIMAL(18,2),
            @HaberLinea2 DECIMAL(18,2),
            @ObservacionDetalle NVARCHAR(300);

        DECLARE cursor_cierre CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            c.TipoCierre,
            c.DescripcionCierre,
            c.PeriodoTrabajo,
            c.MesTrabajo,
            c.IdPlanCuenta,
            c.CodigoCuenta,
            c.NombreCuenta,
            c.CodigoMoneda,
            c.TipoCambioCuenta
        FROM @CuentasProceso AS c
        ORDER BY
            c.TipoCierre ASC,
            c.CodigoCuenta ASC;

        OPEN cursor_cierre;

        FETCH NEXT FROM cursor_cierre
        INTO @TipoCierreTrabajo, @DescripcionCierreTrabajo, @PeriodoTrabajo, @MesTrabajo, @IdPlanCuentaTrabajo,
             @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @CodigoMonedaTrabajo, @TipoCambioCuentaTrabajo;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SELECT
                @SaldoSoles = ISNULL(SUM(
                    CASE
                        WHEN d.DH = 'D'
                            THEN CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Debe END
                        ELSE CASE WHEN d.TotalImporteS > 0 THEN d.TotalImporteS ELSE d.Haber END * -1
                    END), 0),
                @SaldoDolares = ISNULL(SUM(
                    CASE
                        WHEN d.DH = 'D'
                            THEN CASE WHEN d.TotalImporteD > 0 THEN d.TotalImporteD ELSE 0 END
                        ELSE CASE WHEN d.TotalImporteD > 0 THEN d.TotalImporteD ELSE 0 END * -1
                    END), 0)
            FROM dbo.CON_Asiento AS a
            INNER JOIN dbo.CON_AsientoDetalle AS d
                ON d.IdAsiento = a.IdAsiento
            WHERE a.IdEmpresa = @IdEmpresa
              AND a.Periodo BETWEEN @PeriodoBaseInicio AND @PeriodoBaseFin
              AND d.IdPlanCuenta = @IdPlanCuentaTrabajo;

            IF @CodigoMonedaTrabajo = 'USD'
            BEGIN
                SET @TipoCambioAplicado =
                    CASE @TipoCambioCuentaTrabajo
                        WHEN 'C' THEN @TipoCambioCompra
                        WHEN 'V' THEN @TipoCambioVenta
                        ELSE @TipoCambioVenta
                    END;
            END
            ELSE
            BEGIN
                SET @TipoCambioAplicado = 1;
            END;

            IF ABS(@SaldoSoles) < 0.01
            BEGIN
                INSERT INTO dbo.CON_CierreProcesoDetalle
                (
                    IdCierreProceso,
                    TipoCierre,
                    IdPlanCuenta,
                    CodigoMoneda,
                    TipoCambioAplicado,
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
                    @IdCierreProceso,
                    @TipoCierreTrabajo,
                    @IdPlanCuentaTrabajo,
                    @CodigoMonedaTrabajo,
                    @TipoCambioAplicado,
                    NULL,
                    NULL,
                    0,
                    0,
                    N'SIN_SALDO',
                    N'La cuenta no presenta saldo pendiente para el cierre.',
                    @UsuarioRegistro
                );

                FETCH NEXT FROM cursor_cierre
                INTO @TipoCierreTrabajo, @DescripcionCierreTrabajo, @PeriodoTrabajo, @MesTrabajo, @IdPlanCuentaTrabajo,
                     @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @CodigoMonedaTrabajo, @TipoCambioCuentaTrabajo;
                CONTINUE;
            END;

            SET @ImporteAbs = ABS(@SaldoSoles);

            IF @SaldoSoles > 0
            BEGIN
                SET @IdPlanCuentaContra = @IdPlanCuentaPerdida;
                SET @CodigoCuentaContra = @CodigoCuentaPerdida;
                SET @NombreCuentaContra = @NombreCuentaPerdida;
                SET @CodigoMonedaContra = @CodigoMonedaPerdida;
                SET @TipoCambioCuentaContra = @TipoCambioCuentaPerdida;
                SET @DebeLinea1 = 0;
                SET @HaberLinea1 = @ImporteAbs;
                SET @DebeLinea2 = @ImporteAbs;
                SET @HaberLinea2 = 0;
            END
            ELSE
            BEGIN
                SET @IdPlanCuentaContra = @IdPlanCuentaGanancia;
                SET @CodigoCuentaContra = @CodigoCuentaGanancia;
                SET @NombreCuentaContra = @NombreCuentaGanancia;
                SET @CodigoMonedaContra = @CodigoMonedaGanancia;
                SET @TipoCambioCuentaContra = @TipoCambioCuentaGanancia;
                SET @DebeLinea1 = @ImporteAbs;
                SET @HaberLinea1 = 0;
                SET @DebeLinea2 = 0;
                SET @HaberLinea2 = @ImporteAbs;
            END;

            SET @TipoCambioContraAplicado =
                CASE
                    WHEN @CodigoMonedaContra = 'USD' AND @TipoCambioCuentaContra = 'C' THEN @TipoCambioCompra
                    WHEN @CodigoMonedaContra = 'USD' AND @TipoCambioCuentaContra = 'V' THEN @TipoCambioVenta
                    WHEN @CodigoMonedaContra = 'USD' THEN @TipoCambioVenta
                    ELSE 1
                END;

            SET @TotalImporteDLinea1 =
                CASE
                    WHEN @CodigoMonedaTrabajo = 'USD' THEN ABS(@SaldoDolares)
                    ELSE 0
                END;

            SET @TotalImporteDLinea2 =
                CASE
                    WHEN @CodigoMonedaContra = 'USD' AND @TipoCambioContraAplicado > 0 THEN ROUND(@ImporteAbs / @TipoCambioContraAplicado, 2)
                    ELSE 0
                END;

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_CorrelativoAsiento AS c WITH (UPDLOCK, HOLDLOCK)
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @PeriodoTrabajo
            )
            BEGIN
                UPDATE dbo.CON_CorrelativoAsiento
                SET UltimoNumero = UltimoNumero + 1,
                    FechaActualizacion = SYSDATETIME(),
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdOrigen = @IdOrigen
                  AND Periodo = @PeriodoTrabajo;

                SELECT
                    @NumeroAsientoTrabajo = c.UltimoNumero
                FROM dbo.CON_CorrelativoAsiento AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdOrigen = @IdOrigen
                  AND c.Periodo = @PeriodoTrabajo;
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
                    @PeriodoTrabajo,
                    1,
                    SYSDATETIME(),
                    @UsuarioRegistro
                );

                SET @NumeroAsientoTrabajo = 1;
            END;

            SET @GlosaAsiento = CONCAT(@DescripcionCierreTrabajo, N' ', @CodigoCuentaTrabajo, N' - ', @NombreCuentaTrabajo);

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
                @MesTrabajo,
                @PeriodoTrabajo,
                @NumeroAsientoTrabajo,
                @FechaAsiento,
                @FechaAsiento,
                @GlosaAsiento,
                @IdMonedaPen,
                CASE WHEN @CodigoMonedaTrabajo = 'USD' THEN @TipoCambioAplicado ELSE 1 END,
                @ImporteAbs,
                @ImporteAbs,
                N'PROVISIONADO',
                CONCAT(N'CIE-', @PeriodoTrabajo, N'-', @CodigoCuentaTrabajo),
                CONCAT(N'Generado automaticamente por ', @DescripcionCierreTrabajo, N'.'),
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
            VALUES
            (
                @IdAsientoTrabajo,
                1,
                @IdPlanCuentaTrabajo,
                CASE WHEN @DebeLinea1 > 0 THEN 'D' ELSE 'H' END,
                CONCAT(@DescripcionCierreTrabajo, N' / Cuenta origen'),
                @TipoCambioAplicado,
                @DebeLinea1,
                @HaberLinea1,
                @ImporteAbs,
                @TotalImporteDLinea1,
                @UsuarioRegistro
            );

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
            VALUES
            (
                @IdAsientoTrabajo,
                2,
                @IdPlanCuentaContra,
                CASE WHEN @DebeLinea2 > 0 THEN 'D' ELSE 'H' END,
                CONCAT(@DescripcionCierreTrabajo, N' / Contrapartida'),
                @TipoCambioContraAplicado,
                @DebeLinea2,
                @HaberLinea2,
                @ImporteAbs,
                @TotalImporteDLinea2,
                @UsuarioRegistro
            );

            SET @TotalAsientos += 1;
            SET @TotalDebeProceso += @ImporteAbs;
            SET @TotalHaberProceso += @ImporteAbs;
            SET @ObservacionDetalle = CONCAT(N'Asiento generado en periodo ', @PeriodoTrabajo, N' usando cierre por ColBalance.');

            INSERT INTO dbo.CON_CierreProcesoDetalle
            (
                IdCierreProceso,
                TipoCierre,
                IdPlanCuenta,
                CodigoMoneda,
                TipoCambioAplicado,
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
                @IdCierreProceso,
                @TipoCierreTrabajo,
                @IdPlanCuentaTrabajo,
                @CodigoMonedaTrabajo,
                @TipoCambioAplicado,
                @IdAsientoTrabajo,
                @NumeroAsientoTrabajo,
                @ImporteAbs,
                @ImporteAbs,
                N'GENERADO',
                @ObservacionDetalle,
                @UsuarioRegistro
            );

            FETCH NEXT FROM cursor_cierre
            INTO @TipoCierreTrabajo, @DescripcionCierreTrabajo, @PeriodoTrabajo, @MesTrabajo, @IdPlanCuentaTrabajo,
                 @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @CodigoMonedaTrabajo, @TipoCambioCuentaTrabajo;
        END;

        CLOSE cursor_cierre;
        DEALLOCATE cursor_cierre;

        UPDATE dbo.CON_CierreProceso
        SET TotalCuentas = @TotalCuentas,
            TotalAsientos = @TotalAsientos,
            TotalDebe = @TotalDebeProceso,
            TotalHaber = @TotalHaberProceso,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdCierreProceso = @IdCierreProceso;

        COMMIT;

        EXEC dbo.usp_CON_ObtenerCierreProceso
            @IdEmpresa = @IdEmpresa,
            @Anio = @Anio;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

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
