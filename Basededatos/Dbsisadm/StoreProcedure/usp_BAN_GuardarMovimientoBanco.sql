-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Registra o actualiza movimientos de caja y bancos, valida la operacion bancaria, recalcula el importe total, asigna correlativo interno por periodo y guarda referencias documentarias por linea.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Amplia el guardado de Caja y Bancos agregando TipoCambio y Observacion en cabecera, persona por linea, aplicacion de saldos a compras/ventas desde el detalle y validacion obligatoria de cuadre entre Total Operacion y Total Detalle.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Genera y actualiza el asiento contable automatico del movimiento bancario usando el origen configurado para ING y EGR, solo cuando la operacion bancaria tiene indTranConta = 'S'.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Reubica el nro de operacion de la linea bancaria automatica en Referencia para que no se muestre en RUC/DNI dentro del asiento y mantenga consistencia con el detalle del movimiento.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Permite enlazar movimientos de transferencia entre cuentas y reutilizar el guardado bancario desde procesos compuestos sin devolver resultados intermedios.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_GuardarMovimientoBanco
    @IdMovimientoBanco INT = NULL,
    @IdEmpresa INT,
    @IdBancoConfiguracionEmpresa INT,
    @TipoMovimiento CHAR(1),
    @IdOpeBancaria CHAR(2),
    @FechaEmision DATE,
    @TipoCambio DECIMAL(18, 6),
    @IdPersona INT = NULL,
    @NumeroDocumento VARCHAR(20) = NULL,
    @Glosa NVARCHAR(300),
    @Observacion NVARCHAR(500) = NULL,
    @ImporteTotal DECIMAL(18, 2),
    @UsuarioRegistro NVARCHAR(450) = NULL,
    @DetallesXml XML,
    @IdTransferenciaCuenta UNIQUEIDENTIFIER = NULL,
    @RolTransferencia CHAR(1) = NULL,
    @IdMovimientoBancoRelacionado INT = NULL,
    @RetornarResultado BIT = 1,
    @IdMovimientoBancoGenerado INT = NULL OUTPUT,
    @IdAsientoGenerado INT = NULL OUTPUT,
    @NumeroMovimientoGenerado INT = NULL OUTPUT,
    @NumeroAsientoGenerado INT = NULL OUTPUT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @ImporteTotalDebe DECIMAL(18, 2) = 0;
        DECLARE @ImporteTotalHaber DECIMAL(18, 2) = 0;
        DECLARE @NumeroMovimiento INT = 0;
        DECLARE @TotalDetalle DECIMAL(18, 2) = 0;
        DECLARE @TotalDebeAsiento DECIMAL(18, 2) = 0;
        DECLARE @TotalHaberAsiento DECIMAL(18, 2) = 0;
        DECLARE @FechaEmisionOriginal DATE = NULL;
        DECLARE @NumeroMovimientoOriginal INT = NULL;
        DECLARE @IdAsientoTrabajo INT = NULL;
        DECLARE @NumeroAsiento INT = NULL;
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), YEAR(@FechaEmision)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(@FechaEmision)), 2);
        DECLARE @Ejercicio SMALLINT = YEAR(@FechaEmision);
        DECLARE @Mes TINYINT = MONTH(@FechaEmision);
        DECLARE @IdOrigen INT = NULL;
        DECLARE @IdOrigenExistente INT = NULL;
        DECLARE @PeriodoExistente CHAR(6) = NULL;
        DECLARE @EstadoConfiguracion BIT = 0;
        DECLARE @GeneraAsientoAutomatico BIT = 0;
        DECLARE @IndTranConta CHAR(1) = 'N';
        DECLARE @ModuloOperacion CHAR(3) = CASE WHEN @TipoMovimiento = 'I' THEN 'ING' ELSE 'EGR' END;
        DECLARE @IdPlanCuentaBanco INT = NULL;
        DECLARE @IdMoneda INT = NULL;
        DECLARE @NroCuentaCorriente VARCHAR(50) = NULL;
        DECLARE @GlosaAsiento NVARCHAR(500) = NULL;
        DECLARE @ReferenciaExterna NVARCHAR(100) = NULL;

        DECLARE @Detalles TABLE
        (
            Item SMALLINT NOT NULL,
            IdPlanCuenta INT NOT NULL,
            IdPersona INT NULL,
            ModuloOperacionComprobante CHAR(3) NULL,
            IdRegistroComprobante INT NULL,
            ImporteAplicado DECIMAL(18, 2) NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            CodigoCentroCosto VARCHAR(20) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            TipoCambioLinea DECIMAL(18, 6) NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL
        );

        DECLARE @AsientoDetalle TABLE
        (
            Orden INT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NOT NULL,
            CodigoCentroCosto VARCHAR(20) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            TipoCambioLinea DECIMAL(18, 6) NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL
        );

        INSERT INTO @Detalles
        (
            Item,
            IdPlanCuenta,
            IdPersona,
            ModuloOperacionComprobante,
            IdRegistroComprobante,
            ImporteAplicado,
            GlosaDetalle,
            CodigoCentroCosto,
            NumeroDocumento,
            TipoDocumento,
            Serie,
            ReferenciaLinea,
            TipoCambioLinea,
            Debe,
            Haber
        )
        SELECT
            T.N.value('@Item', 'smallint'),
            T.N.value('@IdPlanCuenta', 'int'),
            NULLIF(T.N.value('@IdPersona', 'int'), 0),
            NULLIF(T.N.value('@ModuloOperacionComprobante', 'char(3)'), ''),
            NULLIF(T.N.value('@IdRegistroComprobante', 'int'), 0),
            NULLIF(T.N.value('@ImporteAplicado', 'decimal(18,2)'), 0),
            NULLIF(T.N.value('@GlosaDetalle', 'nvarchar(300)'), N''),
            NULLIF(T.N.value('@CodigoCentroCosto', 'varchar(20)'), ''),
            NULLIF(T.N.value('@NumeroDocumento', 'varchar(20)'), ''),
            NULLIF(T.N.value('@TipoDocumento', 'nvarchar(150)'), N''),
            NULLIF(T.N.value('@Serie', 'varchar(10)'), ''),
            NULLIF(T.N.value('@ReferenciaLinea', 'nvarchar(100)'), N''),
            NULLIF(T.N.value('@TipoCambioLinea', 'decimal(18,6)'), 0),
            T.N.value('@Debe', 'decimal(18,2)'),
            T.N.value('@Haber', 'decimal(18,2)')
        FROM @DetallesXml.nodes('/Detalles/Detalle') AS T(N);

        UPDATE d
        SET ModuloOperacionComprobante = CASE
                                             WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN')
                                              AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                              AND ISNULL(d.ImporteAplicado, 0) > 0
                                                 THEN d.ModuloOperacionComprobante
                                             ELSE NULL
                                         END,
            IdRegistroComprobante = CASE
                                        WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN')
                                         AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                         AND ISNULL(d.ImporteAplicado, 0) > 0
                                            THEN d.IdRegistroComprobante
                                        ELSE NULL
                                    END,
            ImporteAplicado = CASE
                                  WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN')
                                   AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                   AND ISNULL(d.ImporteAplicado, 0) > 0
                                      THEN d.ImporteAplicado
                                  ELSE NULL
                              END
        FROM @Detalles AS d;

        SELECT
            @IdPlanCuentaBanco = cc.IdPlanCuenta,
            @IdMoneda = cc.IdMoneda,
            @NroCuentaCorriente = cc.NroCuentaCorriente
        FROM dbo.CON_BancosConfiguracionEmpresa AS cc
        WHERE cc.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa
          AND cc.IdEmpresa = @IdEmpresa;

        IF @IdPlanCuentaBanco IS NULL
        BEGIN
            RAISERROR('La cuenta corriente seleccionada no pertenece a la empresa.', 16, 1);
        END;

        IF @IdMoneda IS NULL
        BEGIN
            RAISERROR('La cuenta corriente seleccionada no tiene moneda configurada para generar el asiento.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdPlanCuenta = @IdPlanCuentaBanco
              AND pc.IdEmpresa = @IdEmpresa
              AND pc.Estado = 1
              AND pc.AceptaMovimiento = 1
        )
        BEGIN
            RAISERROR('La cuenta contable configurada en la cuenta corriente no es valida para generar el asiento.', 16, 1);
        END;

        IF @TipoMovimiento NOT IN ('I', 'E')
        BEGIN
            RAISERROR('Seleccione un tipo de movimiento valido.', 16, 1);
        END;

        IF @RolTransferencia IS NOT NULL AND @RolTransferencia NOT IN ('E', 'I')
        BEGIN
            RAISERROR('El rol de transferencia debe ser E o I.', 16, 1);
        END;

        IF @TipoCambio <= 0
        BEGIN
            RAISERROR('Ingrese un tipo de cambio mayor a cero.', 16, 1);
        END;

        SELECT TOP (1)
            @IndTranConta = CASE WHEN LTRIM(RTRIM(ISNULL(op.indTranConta, 'N'))) = 'S' THEN 'S' ELSE 'N' END
        FROM dbo.operacionesbancarias AS op
        WHERE LTRIM(RTRIM(op.idOpeBancaria)) = LTRIM(RTRIM(@IdOpeBancaria))
          AND LTRIM(RTRIM(op.Destino)) = @TipoMovimiento;

        IF @IndTranConta NOT IN ('S', 'N')
        BEGIN
            SET @IndTranConta = 'N';
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.operacionesbancarias AS op
            WHERE LTRIM(RTRIM(op.idOpeBancaria)) = LTRIM(RTRIM(@IdOpeBancaria))
              AND LTRIM(RTRIM(op.Destino)) = @TipoMovimiento
        )
        BEGIN
            RAISERROR('La operacion bancaria seleccionada no corresponde al tipo de movimiento.', 16, 1);
        END;

        IF @IndTranConta = 'S'
        BEGIN
            SELECT
                @IdOrigen = c.IdOrigen,
                @EstadoConfiguracion = c.Activo,
                @GeneraAsientoAutomatico = c.GeneraAsientoAutomatico
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.EscenarioOperacion = 'PROVISION'
              AND c.ModuloOperacion = @ModuloOperacion;

            IF @IdOrigen IS NULL
            BEGIN
                RAISERROR('No existe configuracion contable activa para el origen de ingresos o egresos del movimiento bancario.', 16, 1);
            END;

            IF @EstadoConfiguracion = 0
            BEGIN
                RAISERROR('La configuracion contable del movimiento bancario esta inactiva.', 16, 1);
            END;

            IF @GeneraAsientoAutomatico = 0
            BEGIN
                RAISERROR('La configuracion contable del movimiento bancario no esta habilitada para generar asiento automatico.', 16, 1);
            END;
        END;

        IF @IdPersona IS NOT NULL
           AND NOT EXISTS
           (
               SELECT 1
               FROM dbo.ADM_Persona AS p
               WHERE p.IdPersona = @IdPersona
                 AND p.IdEmpresa = @IdEmpresa
           )
        BEGIN
            RAISERROR('La persona vinculada de cabecera no existe o no pertenece a la empresa.', 16, 1);
        END;

        IF NOT EXISTS (SELECT 1 FROM @Detalles)
        BEGIN
            RAISERROR('Debe registrar al menos una linea en el detalle del movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            WHERE d.IdPersona IS NOT NULL
              AND NOT EXISTS
              (
                  SELECT 1
                  FROM dbo.ADM_Persona AS p
                  WHERE p.IdPersona = d.IdPersona
                    AND p.IdEmpresa = @IdEmpresa
              )
        )
        BEGIN
            RAISERROR('Existe una persona en el detalle que no pertenece a la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            LEFT JOIN dbo.CON_PlanCuenta AS pc
                ON pc.IdPlanCuenta = d.IdPlanCuenta
               AND pc.IdEmpresa = @IdEmpresa
               AND pc.Estado = 1
               AND pc.AceptaMovimiento = 1
            WHERE pc.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR('Existe una cuenta contable en el detalle que no pertenece a la empresa, no esta activa o no acepta movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            WHERE d.GlosaDetalle IS NULL
               OR LTRIM(RTRIM(d.GlosaDetalle)) = ''
        )
        BEGIN
            RAISERROR('Cada linea del detalle debe registrar glosa detalle.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            WHERE (d.Debe <= 0 AND d.Haber <= 0)
               OR (d.Debe > 0 AND d.Haber > 0)
        )
        BEGIN
            RAISERROR('Cada linea del detalle debe tener importe solo en Debe o solo en Haber.', 16, 1);
        END;

        SELECT
            @ImporteTotalDebe = ISNULL(SUM(d.Debe), 0),
            @ImporteTotalHaber = ISNULL(SUM(d.Haber), 0)
        FROM @Detalles AS d;

        SET @TotalDetalle = CASE
                                WHEN @TipoMovimiento = 'E' THEN ABS(@ImporteTotalDebe - @ImporteTotalHaber)
                                ELSE ABS(@ImporteTotalHaber - @ImporteTotalDebe)
                            END;

        IF @ImporteTotal <= 0
        BEGIN
            RAISERROR('Ingrese un importe total de cabecera mayor a cero.', 16, 1);
        END;

        IF ABS(@ImporteTotal - @TotalDetalle) >= 0.005
        BEGIN
            RAISERROR('No puede guardar mientras exista diferencia entre Total Operacion y Total Detalle.', 16, 1);
        END;

        BEGIN TRANSACTION;

        IF @IdMovimientoBanco IS NOT NULL
        BEGIN
            SELECT
                @FechaEmisionOriginal = m.FechaEmision,
                @NumeroMovimientoOriginal = m.NumeroMovimiento,
                @IdAsientoTrabajo = m.IdAsiento
            FROM dbo.BAN_MovimientoBanco AS m WITH (UPDLOCK, HOLDLOCK)
            WHERE m.IdMovimientoBanco = @IdMovimientoBanco
              AND m.IdEmpresa = @IdEmpresa;

            IF @FechaEmisionOriginal IS NULL
            BEGIN
                RAISERROR('El movimiento a editar no existe para la empresa activa.', 16, 1);
            END;

            ;WITH AplicacionesPrevias AS
            (
                SELECT
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante,
                    SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
                FROM dbo.BAN_MovimientoBancoDetalle AS d
                WHERE d.IdMovimientoBanco = @IdMovimientoBanco
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN')
                  AND d.IdRegistroComprobante IS NOT NULL
                GROUP BY
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante
            )
            UPDATE c
            SET c.Saldo = CASE
                              WHEN c.Saldo + a.ImporteAplicado > c.ImporteTotal THEN c.ImporteTotal
                              ELSE c.Saldo + a.ImporteAplicado
                          END
            FROM dbo.COM_Compra AS c
            INNER JOIN AplicacionesPrevias AS a
                ON a.ModuloOperacionComprobante = 'COM'
               AND a.IdRegistroComprobante = c.IdCompra
            WHERE c.IdEmpresa = @IdEmpresa;

            ;WITH AplicacionesPrevias AS
            (
                SELECT
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante,
                    SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
                FROM dbo.BAN_MovimientoBancoDetalle AS d
                WHERE d.IdMovimientoBanco = @IdMovimientoBanco
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN')
                  AND d.IdRegistroComprobante IS NOT NULL
                GROUP BY
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante
            )
            UPDATE v
            SET v.Saldo = CASE
                              WHEN v.Saldo + a.ImporteAplicado > v.ImporteTotal THEN v.ImporteTotal
                              ELSE v.Saldo + a.ImporteAplicado
                          END
            FROM dbo.VEN_Venta AS v
            INNER JOIN AplicacionesPrevias AS a
                ON a.ModuloOperacionComprobante = 'VEN'
               AND a.IdRegistroComprobante = v.IdVenta
            WHERE v.IdEmpresa = @IdEmpresa;
        END;

        IF @IdMovimientoBanco IS NULL
           OR YEAR(@FechaEmisionOriginal) <> YEAR(@FechaEmision)
           OR MONTH(@FechaEmisionOriginal) <> MONTH(@FechaEmision)
        BEGIN
            SELECT
                @NumeroMovimiento = ISNULL(MAX(m.NumeroMovimiento), 0) + 1
            FROM dbo.BAN_MovimientoBanco AS m WITH (UPDLOCK, HOLDLOCK)
            WHERE m.IdEmpresa = @IdEmpresa
              AND YEAR(m.FechaEmision) = YEAR(@FechaEmision)
              AND MONTH(m.FechaEmision) = MONTH(@FechaEmision);
        END
        ELSE
        BEGIN
            SET @NumeroMovimiento = ISNULL(@NumeroMovimientoOriginal, 1);
        END;

        SET @GlosaAsiento = LEFT(LTRIM(RTRIM(@Glosa)), 500);
        SET @ReferenciaExterna = LEFT(CONCAT(N'BAN ', @NumeroMovimiento, N' / ', ISNULL(NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''), N'SIN-OPERACION')), 100);

        IF @IndTranConta = 'S'
        BEGIN
            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                GlosaDetalle,
                CodigoCentroCosto,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber
            )
            VALUES
            (
                @IdPlanCuentaBanco,
                LEFT(CONCAT(N'Banco ', ISNULL(@NroCuentaCorriente, N'')), 300),
                NULL,
                NULL,
                NULL,
                NULL,
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                @TipoCambio,
                CASE WHEN @TipoMovimiento = 'I' THEN @ImporteTotal ELSE 0 END,
                CASE WHEN @TipoMovimiento = 'E' THEN @ImporteTotal ELSE 0 END
            );

            INSERT INTO @AsientoDetalle
            (
                IdPlanCuenta,
                GlosaDetalle,
                CodigoCentroCosto,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber
            )
            SELECT
                d.IdPlanCuenta,
                d.GlosaDetalle,
                d.CodigoCentroCosto,
                d.NumeroDocumento,
                d.TipoDocumento,
                d.Serie,
                d.ReferenciaLinea,
                ISNULL(d.TipoCambioLinea, @TipoCambio),
                d.Debe,
                d.Haber
            FROM @Detalles AS d
            ORDER BY d.Item;

            SELECT
                @TotalDebeAsiento = ISNULL(SUM(d.Debe), 0),
                @TotalHaberAsiento = ISNULL(SUM(d.Haber), 0)
            FROM @AsientoDetalle AS d;

            IF @TotalDebeAsiento <> @TotalHaberAsiento
            BEGIN
                RAISERROR('El detalle de Caja y Bancos no genera un asiento contable cuadrado con la cuenta bancaria de cabecera.', 16, 1);
            END;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            LEFT JOIN dbo.COM_Compra AS c
                ON d.ModuloOperacionComprobante = 'COM'
               AND c.IdCompra = d.IdRegistroComprobante
               AND c.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante = 'COM'
              AND c.IdCompra IS NULL
        )
        BEGIN
            RAISERROR('Existe una linea con comprobante de compra que no pertenece a la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            LEFT JOIN dbo.VEN_Venta AS v
                ON d.ModuloOperacionComprobante = 'VEN'
               AND v.IdVenta = d.IdRegistroComprobante
               AND v.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante = 'VEN'
              AND v.IdVenta IS NULL
        )
        BEGIN
            RAISERROR('Existe una linea con comprobante de venta que no pertenece a la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            INNER JOIN dbo.COM_Compra AS c
                ON d.ModuloOperacionComprobante = 'COM'
               AND c.IdCompra = d.IdRegistroComprobante
               AND c.IdEmpresa = @IdEmpresa
            WHERE d.ImporteAplicado > c.Saldo
        )
        BEGIN
            RAISERROR('El importe aplicado en una linea supera el saldo pendiente del comprobante de compra seleccionado.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            INNER JOIN dbo.VEN_Venta AS v
                ON d.ModuloOperacionComprobante = 'VEN'
               AND v.IdVenta = d.IdRegistroComprobante
               AND v.IdEmpresa = @IdEmpresa
            WHERE d.ImporteAplicado > v.Saldo
        )
        BEGIN
            RAISERROR('El importe aplicado en una linea supera el saldo pendiente del comprobante de venta seleccionado.', 16, 1);
        END;

        IF @IndTranConta = 'S' AND @IdAsientoTrabajo IS NULL
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
                @Ejercicio,
                @Mes,
                @Periodo,
                @NumeroAsiento,
                @FechaEmision,
                @FechaEmision,
                @GlosaAsiento,
                @IdMoneda,
                @TipoCambio,
                @TotalDebeAsiento,
                @TotalHaberAsiento,
                N'PROVISIONADO',
                @ReferenciaExterna,
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdAsientoTrabajo = SCOPE_IDENTITY();
        END
        ELSE IF @IndTranConta = 'S'
        BEGIN
            SELECT
                @NumeroAsiento = a.NumeroAsiento,
                @PeriodoExistente = a.Periodo,
                @IdOrigenExistente = a.IdOrigen
            FROM dbo.CON_Asiento AS a
            WHERE a.IdAsiento = @IdAsientoTrabajo
              AND a.IdEmpresa = @IdEmpresa;

            IF @PeriodoExistente IS NULL
            BEGIN
                RAISERROR('El asiento contable vinculado no existe para la empresa activa.', 16, 1);
            END;

            IF @PeriodoExistente <> @Periodo
            BEGIN
                RAISERROR('No se puede cambiar el periodo del movimiento bancario cuando ya tiene asiento contable vinculado.', 16, 1);
            END;

            IF @IdOrigenExistente <> @IdOrigen
            BEGIN
                RAISERROR('No se puede cambiar entre ingreso y egreso cuando el movimiento bancario ya tiene asiento contable vinculado.', 16, 1);
            END;

            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.CON_Asiento
            SET FechaEmision = @FechaEmision,
                FechaAsiento = @FechaEmision,
                Glosa = @GlosaAsiento,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebeAsiento,
                TotalHaber = @TotalHaberAsiento,
                Estado = N'PROVISIONADO',
                ReferenciaExterna = @ReferenciaExterna,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdAsiento = @IdAsientoTrabajo;
        END;
        ELSE IF @IdAsientoTrabajo IS NOT NULL
        BEGIN
            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.BAN_MovimientoBanco
            SET IdAsiento = NULL
            WHERE IdMovimientoBanco = @IdMovimientoBanco
              AND IdEmpresa = @IdEmpresa;

            DELETE FROM dbo.CON_Asiento
            WHERE IdAsiento = @IdAsientoTrabajo
              AND IdEmpresa = @IdEmpresa;

            SET @IdAsientoTrabajo = NULL;
            SET @NumeroAsiento = NULL;
        END;

        IF @IdMovimientoBanco IS NULL
        BEGIN
            INSERT INTO dbo.BAN_MovimientoBanco
            (
                IdEmpresa,
                IdBancoConfiguracionEmpresa,
                TipoMovimiento,
                IdOpeBancaria,
                FechaEmision,
                TipoCambio,
                NumeroMovimiento,
                IdAsiento,
                IdTransferenciaCuenta,
                RolTransferencia,
                IdMovimientoBancoRelacionado,
                IdPersona,
                NumeroDocumento,
                Glosa,
                Observacion,
                ImporteTotal,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdBancoConfiguracionEmpresa,
                @TipoMovimiento,
                @IdOpeBancaria,
                @FechaEmision,
                @TipoCambio,
                @NumeroMovimiento,
                @IdAsientoTrabajo,
                @IdTransferenciaCuenta,
                @RolTransferencia,
                @IdMovimientoBancoRelacionado,
                @IdPersona,
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                LTRIM(RTRIM(@Glosa)),
                NULLIF(LTRIM(RTRIM(@Observacion)), ''),
                @ImporteTotal,
                1,
                NULLIF(LTRIM(RTRIM(@UsuarioRegistro)), '')
            );

            SET @IdMovimientoBanco = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.BAN_MovimientoBanco
            SET IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa,
                TipoMovimiento = @TipoMovimiento,
                IdOpeBancaria = @IdOpeBancaria,
                FechaEmision = @FechaEmision,
                TipoCambio = @TipoCambio,
                NumeroMovimiento = @NumeroMovimiento,
                IdAsiento = @IdAsientoTrabajo,
                IdTransferenciaCuenta = @IdTransferenciaCuenta,
                RolTransferencia = @RolTransferencia,
                IdMovimientoBancoRelacionado = @IdMovimientoBancoRelacionado,
                IdPersona = @IdPersona,
                NumeroDocumento = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                Glosa = LTRIM(RTRIM(@Glosa)),
                Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), ''),
                ImporteTotal = @ImporteTotal,
                UsuarioRegistro = NULLIF(LTRIM(RTRIM(@UsuarioRegistro)), '')
            WHERE IdMovimientoBanco = @IdMovimientoBanco
              AND IdEmpresa = @IdEmpresa;

            DELETE FROM dbo.BAN_MovimientoBancoDetalle
            WHERE IdMovimientoBanco = @IdMovimientoBanco;
        END;

        INSERT INTO dbo.BAN_MovimientoBancoDetalle
        (
            IdMovimientoBanco,
            Item,
            IdPlanCuenta,
            IdPersona,
            ModuloOperacionComprobante,
            IdRegistroComprobante,
            ImporteAplicado,
            GlosaDetalle,
            CodigoCentroCosto,
            NumeroDocumento,
            TipoDocumento,
            Serie,
            ReferenciaLinea,
            TipoCambioLinea,
            Debe,
            Haber
        )
        SELECT
            @IdMovimientoBanco,
            d.Item,
            d.IdPlanCuenta,
            d.IdPersona,
            d.ModuloOperacionComprobante,
            d.IdRegistroComprobante,
            d.ImporteAplicado,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.NumeroDocumento,
            d.TipoDocumento,
            d.Serie,
            d.ReferenciaLinea,
            d.TipoCambioLinea,
            d.Debe,
            d.Haber
        FROM @Detalles AS d
        ORDER BY d.Item;

        IF @IndTranConta = 'S'
        BEGIN
            INSERT INTO dbo.CON_AsientoDetalle
            (
                IdAsiento,
                Item,
                IdPlanCuenta,
                GlosaDetalle,
                CodigoCentroCosto,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber
            )
            SELECT
                @IdAsientoTrabajo,
                ROW_NUMBER() OVER (ORDER BY d.Orden),
                d.IdPlanCuenta,
                d.GlosaDetalle,
                d.CodigoCentroCosto,
                d.TipoDocumento,
                d.NumeroDocumento,
                d.Serie,
                d.ReferenciaLinea,
                ISNULL(d.TipoCambioLinea, @TipoCambio),
                d.Debe,
                d.Haber
            FROM @AsientoDetalle AS d
            ORDER BY d.Orden;
        END;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE c
        SET c.Saldo = c.Saldo - a.ImporteAplicado
        FROM dbo.COM_Compra AS c
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'COM'
           AND a.IdRegistroComprobante = c.IdCompra
        WHERE c.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE v
        SET v.Saldo = v.Saldo - a.ImporteAplicado
        FROM dbo.VEN_Venta AS v
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'VEN'
           AND a.IdRegistroComprobante = v.IdVenta
        WHERE v.IdEmpresa = @IdEmpresa;

        COMMIT TRANSACTION;

        SET @IdMovimientoBancoGenerado = @IdMovimientoBanco;
        SET @IdAsientoGenerado = @IdAsientoTrabajo;
        SET @NumeroMovimientoGenerado = @NumeroMovimiento;
        SET @NumeroAsientoGenerado = @NumeroAsiento;

        IF @RetornarResultado = 1
        BEGIN
            EXEC dbo.usp_BAN_ObtenerMovimientoBanco
                @IdMovimientoBanco = @IdMovimientoBanco,
                @IdEmpresa = @IdEmpresa;
        END;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH

END
