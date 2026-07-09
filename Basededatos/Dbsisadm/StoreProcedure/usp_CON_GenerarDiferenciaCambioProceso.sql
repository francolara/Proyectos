-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/07/2026
-- Description:   Genera o regenera los asientos de diferencia en cambio por cuenta para un periodo.
-- =============================================
-- Firma: FRANCO LARA - 01/07/2026 | Crea un proceso por periodo, elimina la generacion previa del mismo periodo y genera un asiento separado por cuenta usando calculo por saldo o por analisis segun la configuracion del plan de cuentas.
-- Firma: FRANCO LARA - 02/07/2026 | Replica la expansion de cuentas destino y contrapartida para los ajustes de diferencia en cambio cuando la cuenta origen tenga configuracion activa, agrupa el modo analitico por numero de documento, tipo, serie y referencia sin heredar cliente/proveedor al asiento generado, excluye asientos originados por procesos automaticos DIF/AJU/APR/CIE para no recalcular sobre ajustes ya generados y limpia las tablas variables por iteracion para evitar arrastre de analisis entre cuentas.
-- Firma: FRANCO LARA - 03/07/2026 | Usa DH para leer el sentido historico del detalle contable y lo persiste en cada linea generada por diferencia en cambio.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarDiferenciaCambioProceso
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
        DECLARE @UsaTipoCambioSbs BIT = 0
        DECLARE @TipoCambioCompra DECIMAL(18,6)
        DECLARE @TipoCambioVenta DECIMAL(18,6)
        DECLARE @TipoCambioCompraSbs DECIMAL(18,6)
        DECLARE @TipoCambioVentaSbs DECIMAL(18,6)
        DECLARE @CodigoCuentaGanancia VARCHAR(20)
        DECLARE @CodigoCuentaPerdida VARCHAR(20)
        DECLARE @IdPlanCuentaGanancia INT
        DECLARE @IdPlanCuentaPerdida INT
        DECLARE @IdMonedaPen INT
        DECLARE @IdCuentaAdministradora INT
        DECLARE @IdDiferenciaCambioProceso INT
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
          AND c.ModuloOperacion = 'DIF'
          AND c.EscenarioOperacion = 'PROVISION'
          AND c.Activo = 1;

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'No existe una configuracion activa de diferencia en cambio en configuracion contable.', 16, 1);
        END;

        SELECT
            @IdCuentaAdministradora = e.IdCuentaAdministradora
        FROM dbo.SEG_Empresa AS e
        WHERE e.IdEmpresa = @IdEmpresa;

        IF @IdCuentaAdministradora IS NULL
        BEGIN
            RAISERROR(N'La empresa no tiene una cuenta administradora asociada para obtener el tipo de cambio.', 16, 1);
        END;

        SELECT
            @UsaTipoCambioSbs = CASE
                                    WHEN UPPER(LTRIM(RTRIM(pe.ValorParametro))) = 'S'
                                         AND @Mes = 12
                                        THEN 1
                                    ELSE 0
                                END
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'TIPO_CAMBIO_SBS_CIERRE'
          AND pe.Activo = 1;

        SELECT
            @TipoCambioCompra = tc.Compra,
            @TipoCambioVenta = tc.Venta,
            @TipoCambioCompraSbs = tc.CompraSBS,
            @TipoCambioVentaSbs = tc.VentaSBS
        FROM dbo.CON_TipoCambio AS tc
        WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
          AND tc.Fecha = @FechaAsiento
          AND tc.IdMoneda = 'USD'
          AND tc.Estado = 1;

        IF @TipoCambioCompra IS NULL OR @TipoCambioVenta IS NULL
        BEGIN
            RAISERROR(N'No existe tipo de cambio USD para la fecha de cierre del periodo seleccionado.', 16, 1);
        END;

        IF @UsaTipoCambioSbs = 1
        BEGIN
            SET @TipoCambioCompra = @TipoCambioCompraSbs;
            SET @TipoCambioVenta = @TipoCambioVentaSbs;
        END;

        SELECT
            @CodigoCuentaGanancia = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAGANANCIA_DC'
          AND pe.Activo = 1;

        SELECT
            @CodigoCuentaPerdida = NULLIF(LTRIM(RTRIM(pe.ValorParametro)), '')
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'CUENTAPERDIDA_DC'
          AND pe.Activo = 1;

        IF @CodigoCuentaGanancia IS NULL OR @CodigoCuentaPerdida IS NULL
        BEGIN
            RAISERROR(N'Configure las cuentas CUENTAGANANCIA_DC y CUENTAPERDIDA_DC para la empresa activa.', 16, 1);
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
            RAISERROR(N'Las cuentas de ganancia o perdida por diferencia en cambio no existen o no aceptan movimiento en el plan de cuentas.', 16, 1);
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

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND pc.AceptaMovimiento = 1
              AND pc.IdMoneda = 'USD'
              AND ISNULL(pc.TipoCambio, '') NOT IN ('C', 'V')
        )
        BEGIN
            RAISERROR(N'Existen cuentas en dolares activas sin tipo de cambio Compra/Venta configurado.', 16, 1);
        END;

        DECLARE @AsientosEliminar TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        DECLARE @CuentasProceso TABLE
        (
            IdPlanCuenta INT NOT NULL PRIMARY KEY,
            CodigoCuenta VARCHAR(20) NOT NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            TipoCambioCuenta CHAR(1) NOT NULL,
            GeneraPorAnalisis BIT NOT NULL
        );

        INSERT INTO @CuentasProceso
        (
            IdPlanCuenta,
            CodigoCuenta,
            NombreCuenta,
            TipoCambioCuenta,
            GeneraPorAnalisis
        )
        SELECT
            pc.IdPlanCuenta,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.TipoCambio,
            pc.GeneraDiferenciaPorAnalisis
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdEmpresa = @IdEmpresa
          AND pc.Estado = 1
          AND pc.AceptaMovimiento = 1
          AND pc.IdMoneda = 'USD'
          AND pc.TipoCambio IN ('C', 'V');

        SET @TotalCuentas = @@ROWCOUNT;

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;
        BEGIN TRAN;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_DiferenciaCambioProceso AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo
        )
        BEGIN
            INSERT INTO @AsientosEliminar (IdAsiento)
            SELECT DISTINCT
                d.IdAsiento
            FROM dbo.CON_DiferenciaCambioProcesoDetalle AS d
            INNER JOIN dbo.CON_DiferenciaCambioProceso AS p
                ON p.IdDiferenciaCambioProceso = d.IdDiferenciaCambioProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo
              AND d.IdAsiento IS NOT NULL;

            DELETE d
            FROM dbo.CON_DiferenciaCambioProcesoDetalle AS d
            INNER JOIN dbo.CON_DiferenciaCambioProceso AS p
                ON p.IdDiferenciaCambioProceso = d.IdDiferenciaCambioProceso
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.Periodo = @Periodo;

            DELETE p
            FROM dbo.CON_DiferenciaCambioProceso AS p
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

        INSERT INTO dbo.CON_DiferenciaCambioProceso
        (
            IdEmpresa,
            Periodo,
            IdOrigen,
            FechaAsiento,
            UsaTipoCambioSbs,
            TipoCambioCompra,
            TipoCambioVenta,
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
            @UsaTipoCambioSbs,
            @TipoCambioCompra,
            @TipoCambioVenta,
            @TotalCuentas,
            0,
            0,
            0,
            @UsuarioRegistro
        );

        SET @IdDiferenciaCambioProceso = SCOPE_IDENTITY();

        DECLARE
            @IdPlanCuentaTrabajo INT,
            @CodigoCuentaTrabajo VARCHAR(20),
            @NombreCuentaTrabajo NVARCHAR(200),
            @TipoCambioCuenta CHAR(1),
            @GeneraPorAnalisis BIT,
            @TipoCambioAplicado DECIMAL(18,6),
            @DebeGananciaPerdida DECIMAL(18,2),
            @HaberGananciaPerdida DECIMAL(18,2),
            @IdAsientoTrabajo INT,
            @NumeroAsientoTrabajo INT,
            @TotalDebeCuenta DECIMAL(18,2),
            @TotalHaberCuenta DECIMAL(18,2),
            @GlosaAsiento NVARCHAR(500),
            @ObservacionDetalle NVARCHAR(300),
            @AplicoCuentaDestino BIT;

        DECLARE cursor_cuentas CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            c.IdPlanCuenta,
            c.CodigoCuenta,
            c.NombreCuenta,
            c.TipoCambioCuenta,
            c.GeneraPorAnalisis
        FROM @CuentasProceso AS c
        ORDER BY
            c.CodigoCuenta ASC;

        OPEN cursor_cuentas;

        FETCH NEXT FROM cursor_cuentas
        INTO @IdPlanCuentaTrabajo, @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @TipoCambioCuenta, @GeneraPorAnalisis;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            SET @TipoCambioAplicado = CASE WHEN @TipoCambioCuenta = 'C' THEN @TipoCambioCompra ELSE @TipoCambioVenta END;
            SET @DebeGananciaPerdida = 0;
            SET @HaberGananciaPerdida = 0;
            SET @IdAsientoTrabajo = NULL;
            SET @NumeroAsientoTrabajo = NULL;
            SET @ObservacionDetalle = NULL;
            SET @AplicoCuentaDestino = 0;

            DECLARE @MovimientosCuenta TABLE
            (
                IdCliente INT NULL,
                IdProveedor INT NULL,
                NumeroDocumento VARCHAR(20) NULL,
                TipoDocumento NVARCHAR(150) NULL,
                Serie VARCHAR(10) NULL,
                ReferenciaLinea NVARCHAR(100) NULL,
                TotalDebeSoles DECIMAL(18,2) NOT NULL,
                TotalHaberSoles DECIMAL(18,2) NOT NULL,
                TotalSoles DECIMAL(18,2) NOT NULL,
                TotalDebeDolar DECIMAL(18,2) NOT NULL,
                TotalHaberDolar DECIMAL(18,2) NOT NULL,
                TotalDolar DECIMAL(18,2) NOT NULL
            );

            DELETE FROM @MovimientosCuenta;

            IF @GeneraPorAnalisis = 1
            BEGIN
                INSERT INTO @MovimientosCuenta
                (
                    IdCliente,
                    IdProveedor,
                    NumeroDocumento,
                    TipoDocumento,
                    Serie,
                    ReferenciaLinea,
                    TotalDebeSoles,
                    TotalHaberSoles,
                    TotalSoles,
                    TotalDebeDolar,
                    TotalHaberDolar,
                    TotalDolar
                )
                SELECT
                    NULL,
                    NULL,
                    NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''),
                    NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N''),
                    NULLIF(LTRIM(RTRIM(d.Serie)), ''),
                    NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), ''),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END)
                        - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END)
                        - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END)
                FROM dbo.CON_AsientoDetalle AS d
                INNER JOIN dbo.CON_Asiento AS a
                    ON a.IdAsiento = d.IdAsiento
                WHERE a.IdEmpresa = @IdEmpresa
                  AND a.Periodo <= @Periodo
                  AND LEFT(a.Periodo, 4) = LEFT(@Periodo, 4)
                  AND a.IdOrigen <> @IdOrigen
                  AND d.IdPlanCuenta = @IdPlanCuentaTrabajo
                  --AND NOT EXISTS
                  --(
                  --    SELECT 1
                  --    FROM dbo.CON_ConfiguracionContabilizacion AS cfg
                  --    WHERE cfg.IdEmpresa = @IdEmpresa
                  --      AND cfg.IdOrigen = a.IdOrigen
                  --      AND cfg.Activo = 1
                  --      AND cfg.ModuloOperacion IN ('DIF', 'AJU', 'APR', 'CIE')
                  --)
                GROUP BY
                    NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''),
                    NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N''),
                    NULLIF(LTRIM(RTRIM(d.Serie)), ''),
                    NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), '')
                HAVING ABS(
                           SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END)
                           - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END)
                       ) > 0.004
                    OR ABS(
                           SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END)
                           - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END)
                       ) > 0.004;
            END
            ELSE
            BEGIN
                INSERT INTO @MovimientosCuenta
                (
                    IdCliente,
                    IdProveedor,
                    NumeroDocumento,
                    TipoDocumento,
                    Serie,
                    ReferenciaLinea,
                    TotalDebeSoles,
                    TotalHaberSoles,
                    TotalSoles,
                    TotalDebeDolar,
                    TotalHaberDolar,
                    TotalDolar
                )
                SELECT
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END)
                        - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END),
                    SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END)
                        - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END)
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
                HAVING ABS(
                           SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteD ELSE 0 END)
                           - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteD ELSE 0 END)
                       ) > 0.004
                    OR ABS(
                           SUM(CASE WHEN d.DH = 'D' THEN d.TotalImporteS ELSE 0 END)
                           - SUM(CASE WHEN d.DH = 'H' THEN d.TotalImporteS ELSE 0 END)
                       ) > 0.004;
            END;

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
                LEFT(CONCAT(N'Ajuste T.C. ', FORMAT(@TipoCambioAplicado, '0.0000'), N' por ', FORMAT(m.TotalDolar, '0.00'), N' USD'), 300),
                m.NumeroDocumento,
                m.TipoDocumento,
                m.Serie,
                m.ReferenciaLinea,
                NULL,
                NULL,
                @TipoCambioAplicado,
                CASE WHEN m.TotalSoles < ROUND(m.TotalDolar * @TipoCambioAplicado, 2) THEN ROUND((m.TotalDolar * @TipoCambioAplicado) - m.TotalSoles, 2) ELSE 0 END,
                CASE WHEN m.TotalSoles > ROUND(m.TotalDolar * @TipoCambioAplicado, 2) THEN ROUND(m.TotalSoles - (m.TotalDolar * @TipoCambioAplicado), 2) ELSE 0 END,
                CASE
                    WHEN m.TotalSoles < ROUND(m.TotalDolar * @TipoCambioAplicado, 2) THEN ROUND((m.TotalDolar * @TipoCambioAplicado) - m.TotalSoles, 2)
                    WHEN m.TotalSoles > ROUND(m.TotalDolar * @TipoCambioAplicado, 2) THEN ROUND(m.TotalSoles - (m.TotalDolar * @TipoCambioAplicado), 2)
                    ELSE 0
                END,
                0
            FROM @MovimientosCuenta AS m
            WHERE ROUND(m.TotalSoles, 2) <> ROUND(m.TotalDolar * @TipoCambioAplicado, 2);

            SELECT
                @HaberGananciaPerdida = ISNULL(SUM(l.Debe), 0),
                @DebeGananciaPerdida = ISNULL(SUM(l.Haber), 0)
            FROM @LineasCuenta AS l;

            IF @DebeGananciaPerdida > 0
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
                    N'PERDIDA POR DIF. CAMBIO',
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    @TipoCambioAplicado,
                    @DebeGananciaPerdida,
                    0,
                    @DebeGananciaPerdida,
                    0
                );
            END;

            IF @HaberGananciaPerdida > 0
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
                    N'GANANCIA POR DIF. CAMBIO',
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    NULL,
                    @TipoCambioAplicado,
                    0,
                    @HaberGananciaPerdida,
                    @HaberGananciaPerdida,
                    0
                );
            END;

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
                DECLARE @IdCuentaCargoDestino INT
                DECLARE @IdCuentaAbonoDestino INT
                DECLARE @PorcentajeDestino DECIMAL(7,4)
                DECLARE @EsUltimoDestino BIT
                DECLARE @ImporteBaseDestino DECIMAL(18,2)
                DECLARE @ImporteDistribuidoDestino DECIMAL(18,2)
                DECLARE @ImporteTramoDestino DECIMAL(18,2)

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
                    l.Haber
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
                     @DebeOrigenDestino, @HaberOrigenDestino;

                WHILE @@FETCH_STATUS = 0
                BEGIN
                    SET @ImporteBaseDestino = CASE
                                                  WHEN @DebeOrigenDestino > 0 THEN @DebeOrigenDestino
                                                  ELSE @HaberOrigenDestino
                                              END;
                    SET @ImporteDistribuidoDestino = 0;

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
                        SET @ImporteTramoDestino = CASE
                                                       WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDestino - @ImporteDistribuidoDestino
                                                       ELSE ROUND(@ImporteBaseDestino * (@PorcentajeDestino / 100.0), 2)
                                                   END;

                        IF @ImporteTramoDestino <> 0
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
                                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'Ajuste diferencia en cambio'), N' / Destino'), 300),
                                @NumeroDocumentoDestino,
                                @TipoDocumentoDestino,
                                @SerieDestino,
                                @ReferenciaLineaDestino,
                                NULL,
                                NULL,
                                @TipoCambioLineaDestino,
                                @ImporteTramoDestino,
                                0,
                                @ImporteTramoDestino,
                                0
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
                                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'Ajuste diferencia en cambio'), N' / Contrapartida'), 300),
                                @NumeroDocumentoDestino,
                                @TipoDocumentoDestino,
                                @SerieDestino,
                                @ReferenciaLineaDestino,
                                NULL,
                                NULL,
                                @TipoCambioLineaDestino,
                                0,
                                @ImporteTramoDestino,
                                @ImporteTramoDestino,
                                0
                            );

                            SET @AplicoCuentaDestino = 1;
                        END;

                        SET @ImporteDistribuidoDestino = @ImporteDistribuidoDestino + @ImporteTramoDestino;

                        FETCH NEXT FROM cursor_tramo_destino
                        INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
                    END;

                    CLOSE cursor_tramo_destino;
                    DEALLOCATE cursor_tramo_destino;

                    FETCH NEXT FROM cursor_linea_destino
                    INTO @IdPlanCuentaOrigenDestino, @ItemLineaDestino, @GlosaBaseDestino, @NumeroDocumentoDestino, @TipoDocumentoDestino, @SerieDestino,
                         @ReferenciaLineaDestino, @IdClienteDestino, @IdProveedorDestino, @TipoCambioLineaDestino,
                         @DebeOrigenDestino, @HaberOrigenDestino;
                END;

                CLOSE cursor_linea_destino;
                DEALLOCATE cursor_linea_destino;
            END;

            SELECT
                @TotalDebeCuenta = ISNULL(SUM(l.Debe), 0),
                @TotalHaberCuenta = ISNULL(SUM(l.Haber), 0)
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

                SET @GlosaAsiento = CONCAT(N'DIFERENCIA DE CAMBIO ', @CodigoCuentaTrabajo, N' - ', @NombreCuentaTrabajo);

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
                    @IdMonedaPen,
                    @TipoCambioAplicado,
                    @TotalDebeCuenta,
                    @TotalHaberCuenta,
                    N'PROVISIONADO',
                    CONCAT(N'DIFC-', @Periodo, N'-', @CodigoCuentaTrabajo),
                    CASE
                        WHEN @GeneraPorAnalisis = 1 AND @AplicoCuentaDestino = 1 THEN N'Generado por analisis con cuentas destino'
                        WHEN @GeneraPorAnalisis = 1 THEN N'Generado por analisis'
                        WHEN @AplicoCuentaDestino = 1 THEN N'Generado por saldo con cuentas destino'
                        ELSE N'Generado por saldo'
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
                    l.Item,
                    l.IdPlanCuenta,
                    CASE WHEN l.Debe > 0 THEN 'D' ELSE 'H' END,
                    l.GlosaDetalle,
                    l.NumeroDocumento,
                    l.TipoDocumento,
                    l.Serie,
                    l.ReferenciaLinea,
                    l.TipoCambioLinea,
                    l.IdCliente,
                    l.IdProveedor,
                    l.Debe,
                    l.Haber,
                    l.TotalImporteS,
                    l.TotalImporteD,
                    @UsuarioRegistro
                FROM @LineasCuenta AS l
                ORDER BY
                    l.Item ASC;

                SET @TotalAsientos += 1;
                SET @TotalDebeProceso += @TotalDebeCuenta;
                SET @TotalHaberProceso += @TotalHaberCuenta;
                SET @ObservacionDetalle = CASE
                                              WHEN @GeneraPorAnalisis = 1 AND @AplicoCuentaDestino = 1 THEN N'Asiento generado por analisis documental con expansion de cuentas destino.'
                                              WHEN @GeneraPorAnalisis = 1 THEN N'Asiento generado por analisis documental.'
                                              WHEN @AplicoCuentaDestino = 1 THEN N'Asiento generado por saldo consolidado con expansion de cuentas destino.'
                                              ELSE N'Asiento generado por saldo consolidado.'
                                          END;

                INSERT INTO dbo.CON_DiferenciaCambioProcesoDetalle
                (
                    IdDiferenciaCambioProceso,
                    IdPlanCuenta,
                    GeneraPorAnalisis,
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
                    @IdDiferenciaCambioProceso,
                    @IdPlanCuentaTrabajo,
                    @GeneraPorAnalisis,
                    @TipoCambioAplicado,
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
                INSERT INTO dbo.CON_DiferenciaCambioProcesoDetalle
                (
                    IdDiferenciaCambioProceso,
                    IdPlanCuenta,
                    GeneraPorAnalisis,
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
                    @IdDiferenciaCambioProceso,
                    @IdPlanCuentaTrabajo,
                    @GeneraPorAnalisis,
                    @TipoCambioAplicado,
                    NULL,
                    NULL,
                    0,
                    0,
                    N'SIN_DIFERENCIA',
                    N'La cuenta no genero diferencia en cambio para el periodo seleccionado.',
                    @UsuarioRegistro
                );
            END;

            FETCH NEXT FROM cursor_cuentas
            INTO @IdPlanCuentaTrabajo, @CodigoCuentaTrabajo, @NombreCuentaTrabajo, @TipoCambioCuenta, @GeneraPorAnalisis;
        END;

        CLOSE cursor_cuentas;
        DEALLOCATE cursor_cuentas;

        UPDATE dbo.CON_DiferenciaCambioProceso
        SET TotalCuentas = @TotalCuentas,
            TotalAsientos = @TotalAsientos,
            TotalDebe = @TotalDebeProceso,
            TotalHaber = @TotalHaberProceso,
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdDiferenciaCambioProceso = @IdDiferenciaCambioProceso;

        COMMIT;
        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        EXEC dbo.usp_CON_ObtenerDiferenciaCambioProceso
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
