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
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Conserva las lineas originales del detalle y agrega cuentas destino y contrapartida segun la configuracion activa por cuenta contable.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Permite pagar documentos de detraccion y percepcion desde Caja y Bancos enlazando saldos pendientes de compras.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Corrige documentos enlazados para conservar tipo documento por codigo, recalcula importes por moneda y soporta percepciones de compras como pendiente independiente.
-- Firma: FRANCO LARA - 29/06/2026 | Vuelve obligatorio el tipo de cambio por linea en Caja y Bancos y en el asiento automatico generado desde su detalle.
-- Firma: FRANCO LARA - 30/06/2026 | Agrega soporte de pagos para documentos pendientes de Renta4ta (R4T) originados desde compras de recibos por honorarios.
-- Firma: FRANCO LARA - 03/07/2026 | Valida el sentido contable del detalle bancario segun el tipo de movimiento para evitar ingresos con exceso de Debe o egresos con exceso de Haber que cuadraban visualmente por valor absoluto pero desbalanceaban el asiento automatico; ademas persiste el Periodo del movimiento bancario desde FechaEmision para mantener consistente el listado operativo con la fecha real grabada y ahora guarda DH en cada linea del asiento automatico.
-- Firma: FRANCO LARA - 04/07/2026 | Convierte los pagos de Caja y Bancos a la moneda del comprobante antes de afectar saldos, topa la aplicacion para no dejar compras/ventas/documentos auxiliares en negativo, guarda en ImporteAplicado solo el monto efectivamente consumido por cada saldo documentario y resuelve la moneda documental desde ADM_Moneda para esquemas donde compras/ventas guardan IdMoneda.
-- Firma: FRANCO LARA - 06/07/2026 | Cuando un documento queda cancelado al 100 por ciento, agrega lineas analiticas de ajuste cambiario en soles y/o dolares usando las cuentas de ganancia/perdida de diferencia en cambio sin alterar el Debe/Haber del asiento bancario.

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
        DECLARE @CodigoMonedaCuenta VARCHAR(10) = NULL;
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
            TipoCambioLinea DECIMAL(18, 6) NOT NULL,
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
            TipoCambioLinea DECIMAL(18, 6) NOT NULL,
            Debe DECIMAL(18, 2) NOT NULL,
            Haber DECIMAL(18, 2) NOT NULL
        );

        DECLARE @DetalleDestinoBase TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            CodigoCentroCosto VARCHAR(20) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            TipoCambioLinea DECIMAL(18,6) NOT NULL,
            ImporteBase DECIMAL(18,2) NOT NULL
        );

        DECLARE @CuentaDestinoDetalle TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            Orden SMALLINT NOT NULL,
            IdPlanCuentaDestinoCargo INT NOT NULL,
            IdPlanCuentaDestinoAbono INT NOT NULL,
            Porcentaje DECIMAL(7,4) NOT NULL,
            EsUltimo BIT NOT NULL
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
            T.N.value('@TipoCambioLinea', 'decimal(18,6)'),
            T.N.value('@Debe', 'decimal(18,2)'),
            T.N.value('@Haber', 'decimal(18,2)')
        FROM @DetallesXml.nodes('/Detalles/Detalle') AS T(N);

        UPDATE d
        SET ModuloOperacionComprobante = CASE
                                             WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                                              AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                              AND ISNULL(d.ImporteAplicado, 0) > 0
                                                 THEN d.ModuloOperacionComprobante
                                             ELSE NULL
                                         END,
            IdRegistroComprobante = CASE
                                        WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                                         AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                         AND ISNULL(d.ImporteAplicado, 0) > 0
                                            THEN d.IdRegistroComprobante
                                        ELSE NULL
                                    END,
            ImporteAplicado = CASE
                                  WHEN d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                                   AND ISNULL(d.IdRegistroComprobante, 0) > 0
                                   AND ISNULL(d.ImporteAplicado, 0) > 0
                                      THEN d.ImporteAplicado
                                  ELSE NULL
                              END
        FROM @Detalles AS d;

        SELECT
            @IdPlanCuentaBanco = cc.IdPlanCuenta,
            @IdMoneda = cc.IdMoneda,
            @CodigoMonedaCuenta = m.CodigoMoneda,
            @NroCuentaCorriente = cc.NroCuentaCorriente
        FROM dbo.CON_BancosConfiguracionEmpresa AS cc
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = cc.IdMoneda
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

        SET @CodigoMonedaCuenta = UPPER(LTRIM(RTRIM(ISNULL(@CodigoMonedaCuenta, ''))));

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

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            WHERE d.TipoCambioLinea <= 0
        )
        BEGIN
            RAISERROR('Cada linea del detalle debe tener tipo de cambio mayor a cero.', 16, 1);
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

        IF @TipoMovimiento = 'I' AND @ImporteTotalHaber <= @ImporteTotalDebe
        BEGIN
            RAISERROR('En ingresos, el detalle debe tener mayor Haber que Debe para compensar la cuenta bancaria.', 16, 1);
        END;

        IF @TipoMovimiento = 'E' AND @ImporteTotalDebe <= @ImporteTotalHaber
        BEGIN
            RAISERROR('En egresos, el detalle debe tener mayor Debe que Haber para compensar la cuenta bancaria.', 16, 1);
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
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
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
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                  AND d.IdRegistroComprobante IS NOT NULL
                GROUP BY
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante
            )
            UPDATE cd
            SET cd.Saldo = CASE
                               WHEN cd.Saldo + a.ImporteAplicado > cd.ImporteDetraccion THEN cd.ImporteDetraccion
                               ELSE cd.Saldo + a.ImporteAplicado
                           END
            FROM dbo.COM_CompraDetraccion AS cd
            INNER JOIN AplicacionesPrevias AS a
                ON a.ModuloOperacionComprobante = 'DET'
               AND a.IdRegistroComprobante = cd.IdCompraDetraccion
            WHERE cd.IdEmpresa = @IdEmpresa;

            ;WITH AplicacionesPrevias AS
            (
                SELECT
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante,
                    SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
                FROM dbo.BAN_MovimientoBancoDetalle AS d
                WHERE d.IdMovimientoBanco = @IdMovimientoBanco
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                  AND d.IdRegistroComprobante IS NOT NULL
                GROUP BY
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante
            )
            UPDATE cp
            SET cp.Saldo = CASE
                               WHEN cp.Saldo + a.ImporteAplicado > cp.ImportePercepcion THEN cp.ImportePercepcion
                               ELSE cp.Saldo + a.ImporteAplicado
                           END
            FROM dbo.COM_CompraPercepcion AS cp
            INNER JOIN AplicacionesPrevias AS a
                ON a.ModuloOperacionComprobante = 'PER'
               AND a.IdRegistroComprobante = cp.IdCompraPercepcion
            WHERE cp.IdEmpresa = @IdEmpresa;

            ;WITH AplicacionesPrevias AS
            (
                SELECT
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante,
                    SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
                FROM dbo.BAN_MovimientoBancoDetalle AS d
                WHERE d.IdMovimientoBanco = @IdMovimientoBanco
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
                  AND d.IdRegistroComprobante IS NOT NULL
                GROUP BY
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante
            )
            UPDATE cr
            SET cr.Saldo = CASE
                               WHEN cr.Saldo + a.ImporteAplicado > cr.Retencion THEN cr.Retencion
                               ELSE cr.Saldo + a.ImporteAplicado
                           END
            FROM dbo.COM_CompraRetencion AS cr
            INNER JOIN AplicacionesPrevias AS a
                ON a.ModuloOperacionComprobante = 'R4T'
               AND a.IdRegistroComprobante = cr.IdCompraRetencion
            WHERE cr.IdEmpresa = @IdEmpresa;

            ;WITH AplicacionesPrevias AS
            (
                SELECT
                    d.ModuloOperacionComprobante,
                    d.IdRegistroComprobante,
                    SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
                FROM dbo.BAN_MovimientoBancoDetalle AS d
                WHERE d.IdMovimientoBanco = @IdMovimientoBanco
                  AND d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
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
                d.TipoCambioLinea,
                d.Debe,
                d.Haber
            FROM @Detalles AS d
            ORDER BY d.Item;

            INSERT INTO @DetalleDestinoBase
            (
                IdPlanCuentaOrigen,
                GlosaDetalle,
                CodigoCentroCosto,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                ImporteBase
            )
            SELECT
                d.IdPlanCuenta,
                d.GlosaDetalle,
                d.CodigoCentroCosto,
                d.NumeroDocumento,
                d.TipoDocumento,
                d.Serie,
                d.ReferenciaLinea,
                d.TipoCambioLinea,
                CASE
                    WHEN d.Debe > 0 THEN d.Debe
                    ELSE d.Haber
                END
            FROM @Detalles AS d
            WHERE (d.Debe > 0 OR d.Haber > 0);

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
                    b.IdPlanCuentaOrigen
                FROM @DetalleDestinoBase AS b
            ) AS base
                ON base.IdPlanCuentaOrigen = r.IdPlanCuentaOrigen
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
                RAISERROR('Existe una configuracion activa de cuentas destino con cuentas cargo o abono invalidas para la empresa.', 16, 1);
            END;

            DECLARE @IdPlanCuentaOrigenDestino INT;
            DECLARE @GlosaBaseDestino NVARCHAR(300);
            DECLARE @CodigoCentroCostoDestino VARCHAR(20);
            DECLARE @NumeroDocumentoDestino VARCHAR(20);
            DECLARE @TipoDocumentoDestino NVARCHAR(150);
            DECLARE @SerieDestino VARCHAR(10);
            DECLARE @ReferenciaDestino NVARCHAR(100);
            DECLARE @TipoCambioDestino DECIMAL(18,6);
            DECLARE @ImporteBaseDestino DECIMAL(18,2);
            DECLARE @IdCuentaCargoDestino INT;
            DECLARE @IdCuentaAbonoDestino INT;
            DECLARE @PorcentajeDestino DECIMAL(7,4);
            DECLARE @EsUltimoDestino BIT;
            DECLARE @ImporteDistribuidoDestino DECIMAL(18,2);
            DECLARE @ImporteTramoDestino DECIMAL(18,2);

            DECLARE cursor_linea_destino CURSOR LOCAL FAST_FORWARD FOR
            SELECT
                b.IdPlanCuentaOrigen,
                b.GlosaDetalle,
                b.CodigoCentroCosto,
                b.NumeroDocumento,
                b.TipoDocumento,
                b.Serie,
                b.ReferenciaLinea,
                b.TipoCambioLinea,
                b.ImporteBase
            FROM @DetalleDestinoBase AS b
            WHERE b.ImporteBase > 0
              AND EXISTS
              (
                  SELECT 1
                  FROM @CuentaDestinoDetalle AS d
                  WHERE d.IdPlanCuentaOrigen = b.IdPlanCuentaOrigen
              );

            OPEN cursor_linea_destino;

            FETCH NEXT FROM cursor_linea_destino
            INTO @IdPlanCuentaOrigenDestino, @GlosaBaseDestino, @CodigoCentroCostoDestino, @NumeroDocumentoDestino,
                 @TipoDocumentoDestino, @SerieDestino, @ReferenciaDestino, @TipoCambioDestino, @ImporteBaseDestino;

            WHILE @@FETCH_STATUS = 0
            BEGIN
                SET @ImporteDistribuidoDestino = 0;

                DECLARE cursor_tramo_destino CURSOR LOCAL FAST_FORWARD FOR
                SELECT
                    d.IdPlanCuentaDestinoCargo,
                    d.IdPlanCuentaDestinoAbono,
                    d.Porcentaje,
                    d.EsUltimo
                FROM @CuentaDestinoDetalle AS d
                WHERE d.IdPlanCuentaOrigen = @IdPlanCuentaOrigenDestino
                ORDER BY
                    d.Orden ASC;

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
                            @IdCuentaCargoDestino,
                            LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), ''), N'Distribucion cuenta destino'), N' / Destino'), 300),
                            @CodigoCentroCostoDestino,
                            @NumeroDocumentoDestino,
                            @TipoDocumentoDestino,
                            @SerieDestino,
                            @ReferenciaDestino,
                            @TipoCambioDestino,
                            @ImporteTramoDestino,
                            0
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
                        VALUES
                        (
                            @IdCuentaAbonoDestino,
                            LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), ''), N'Distribucion cuenta destino'), N' / Contrapartida'), 300),
                            @CodigoCentroCostoDestino,
                            @NumeroDocumentoDestino,
                            @TipoDocumentoDestino,
                            @SerieDestino,
                            @ReferenciaDestino,
                            @TipoCambioDestino,
                            0,
                            @ImporteTramoDestino
                        );
                    END;

                    SET @ImporteDistribuidoDestino = @ImporteDistribuidoDestino + @ImporteTramoDestino;

                    FETCH NEXT FROM cursor_tramo_destino
                    INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
                END;

                CLOSE cursor_tramo_destino;
                DEALLOCATE cursor_tramo_destino;

                FETCH NEXT FROM cursor_linea_destino
                INTO @IdPlanCuentaOrigenDestino, @GlosaBaseDestino, @CodigoCentroCostoDestino, @NumeroDocumentoDestino,
                     @TipoDocumentoDestino, @SerieDestino, @ReferenciaDestino, @TipoCambioDestino, @ImporteBaseDestino;
            END;

            CLOSE cursor_linea_destino;
            DEALLOCATE cursor_linea_destino;

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
            LEFT JOIN dbo.COM_CompraDetraccion AS cd
                ON d.ModuloOperacionComprobante = 'DET'
               AND cd.IdCompraDetraccion = d.IdRegistroComprobante
               AND cd.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante = 'DET'
              AND cd.IdCompraDetraccion IS NULL
        )
        BEGIN
            RAISERROR('Existe una linea con documento de detraccion que no pertenece a la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            LEFT JOIN dbo.COM_CompraPercepcion AS cp
                ON d.ModuloOperacionComprobante = 'PER'
               AND cp.IdCompraPercepcion = d.IdRegistroComprobante
               AND cp.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante = 'PER'
              AND cp.IdCompraPercepcion IS NULL
        )
        BEGIN
            RAISERROR('Existe una linea con documento de percepcion que no pertenece a la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalles AS d
            LEFT JOIN dbo.COM_CompraRetencion AS cr
                ON d.ModuloOperacionComprobante = 'R4T'
               AND cr.IdCompraRetencion = d.IdRegistroComprobante
               AND cr.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante = 'R4T'
              AND cr.IdCompraRetencion IS NULL
        )
        BEGIN
            RAISERROR('Existe una linea con documento de Renta4ta que no pertenece a la empresa activa.', 16, 1);
        END;

        ;WITH AplicacionesSolicitadas AS
        (
            SELECT
                d.Item,
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SaldoDocumento = COALESCE(c.Saldo, v.Saldo, cd.Saldo, cp.Saldo, cr.Saldo, 0),
                ImporteDocumentoSolicitado = CASE
                                                 WHEN @CodigoMonedaCuenta = UPPER(LTRIM(RTRIM(COALESCE(mc.CodigoMoneda, mv.CodigoMoneda, mcd.CodigoMoneda, mcp.CodigoMoneda, mcr.CodigoMoneda, @CodigoMonedaCuenta))))
                                                     THEN ISNULL(d.ImporteAplicado, 0)
                                                 WHEN @CodigoMonedaCuenta = 'PEN'
                                                  AND UPPER(LTRIM(RTRIM(COALESCE(mc.CodigoMoneda, mv.CodigoMoneda, mcd.CodigoMoneda, mcp.CodigoMoneda, mcr.CodigoMoneda, '')))) = 'USD'
                                                     THEN ROUND(ISNULL(d.ImporteAplicado, 0) / NULLIF(d.TipoCambioLinea, 0), 2)
                                                 WHEN @CodigoMonedaCuenta = 'USD'
                                                  AND UPPER(LTRIM(RTRIM(COALESCE(mc.CodigoMoneda, mv.CodigoMoneda, mcd.CodigoMoneda, mcp.CodigoMoneda, mcr.CodigoMoneda, '')))) = 'PEN'
                                                     THEN ROUND(ISNULL(d.ImporteAplicado, 0) * d.TipoCambioLinea, 2)
                                                 ELSE ISNULL(d.ImporteAplicado, 0)
                                             END
            FROM @Detalles AS d
            LEFT JOIN dbo.COM_Compra AS c
                ON d.ModuloOperacionComprobante = 'COM'
               AND c.IdCompra = d.IdRegistroComprobante
               AND c.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.ADM_Moneda AS mc
                ON mc.IdMoneda = c.IdMoneda
            LEFT JOIN dbo.VEN_Venta AS v
                ON d.ModuloOperacionComprobante = 'VEN'
               AND v.IdVenta = d.IdRegistroComprobante
               AND v.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.ADM_Moneda AS mv
                ON mv.IdMoneda = v.IdMoneda
            LEFT JOIN dbo.COM_CompraDetraccion AS cd
                ON d.ModuloOperacionComprobante = 'DET'
               AND cd.IdCompraDetraccion = d.IdRegistroComprobante
               AND cd.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.ADM_Moneda AS mcd
                ON mcd.IdMoneda = cd.IdMoneda
            LEFT JOIN dbo.COM_CompraPercepcion AS cp
                ON d.ModuloOperacionComprobante = 'PER'
               AND cp.IdCompraPercepcion = d.IdRegistroComprobante
               AND cp.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.ADM_Moneda AS mcp
                ON mcp.IdMoneda = cp.IdMoneda
            LEFT JOIN dbo.COM_CompraRetencion AS cr
                ON d.ModuloOperacionComprobante = 'R4T'
               AND cr.IdCompraRetencion = d.IdRegistroComprobante
               AND cr.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.ADM_Moneda AS mcr
                ON mcr.IdMoneda = cr.IdMoneda
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
              AND ISNULL(d.ImporteAplicado, 0) > 0
        ),
        AplicacionesDistribuidas AS
        (
            SELECT
                a.Item,
                a.SaldoDocumento,
                a.ImporteDocumentoSolicitado,
                ImporteDocumentoPrevio = ISNULL(
                    SUM(a.ImporteDocumentoSolicitado) OVER (
                        PARTITION BY a.ModuloOperacionComprobante, a.IdRegistroComprobante
                        ORDER BY a.Item
                        ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING), 0)
            FROM AplicacionesSolicitadas AS a
        )
        UPDATE d
        SET d.ImporteAplicado = CASE
                                    WHEN a.SaldoDocumento <= a.ImporteDocumentoPrevio THEN 0
                                    WHEN a.ImporteDocumentoSolicitado <= a.SaldoDocumento - a.ImporteDocumentoPrevio THEN a.ImporteDocumentoSolicitado
                                    ELSE a.SaldoDocumento - a.ImporteDocumentoPrevio
                                END
        FROM @Detalles AS d
        INNER JOIN AplicacionesDistribuidas AS a
            ON a.Item = d.Item;

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
                Periodo,
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
                @Periodo,
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
                Periodo = @Periodo,
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
            Haber,
            TotalImporteS,
            TotalImporteD
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
            calc.TipoCambioAplicado,
            d.Debe,
            d.Haber,
            CASE
                WHEN @CodigoMonedaCuenta = 'USD' THEN ROUND(calc.ImporteLinea * calc.TipoCambioAplicado, 2)
                ELSE calc.ImporteLinea
            END,
            CASE
                WHEN @CodigoMonedaCuenta = 'USD' THEN calc.ImporteLinea
                ELSE ROUND(calc.ImporteLinea / NULLIF(calc.TipoCambioAplicado, 0), 2)
            END
        FROM @Detalles AS d
        CROSS APPLY
        (
            SELECT
                CASE
                    WHEN d.Debe > 0 THEN d.Debe
                    ELSE d.Haber
                END AS ImporteLinea,
                d.TipoCambioLinea AS TipoCambioAplicado
        ) AS calc
        ORDER BY d.Item;

        IF @IndTranConta = 'S'
        BEGIN
            INSERT INTO dbo.CON_AsientoDetalle
            (
                IdAsiento,
                Item,
                IdPlanCuenta,
                DH,
                GlosaDetalle,
                CodigoCentroCosto,
                TipoDocumento,
                NumeroDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD
            )
            SELECT
                @IdAsientoTrabajo,
                ROW_NUMBER() OVER (ORDER BY d.Orden),
                d.IdPlanCuenta,
                calc.Dh,
                d.GlosaDetalle,
                d.CodigoCentroCosto,
                d.TipoDocumento,
                d.NumeroDocumento,
                d.Serie,
                d.ReferenciaLinea,
                calc.TipoCambioAplicado,
                d.Debe,
                d.Haber,
                CASE
                    WHEN @CodigoMonedaCuenta = 'USD' THEN ROUND(calc.ImporteLinea * calc.TipoCambioAplicado, 2)
                    ELSE calc.ImporteLinea
                END,
                CASE
                    WHEN @CodigoMonedaCuenta = 'USD' THEN calc.ImporteLinea
                    ELSE ROUND(calc.ImporteLinea / NULLIF(calc.TipoCambioAplicado, 0), 2)
                END
            FROM @AsientoDetalle AS d
            CROSS APPLY
            (
                SELECT
                    CASE
                        WHEN d.Debe > 0 THEN d.Debe
                        ELSE d.Haber
                    END AS ImporteLinea,
                    CASE WHEN d.Debe > 0 THEN 'D' ELSE 'H' END AS Dh,
                    d.TipoCambioLinea AS TipoCambioAplicado
            ) AS calc
            ORDER BY d.Orden;
        END;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE c
        SET c.Saldo = CASE
                          WHEN c.Saldo - a.ImporteAplicado < 0 THEN 0
                          ELSE c.Saldo - a.ImporteAplicado
                      END
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
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE cd
        SET cd.Saldo = CASE
                           WHEN cd.Saldo - a.ImporteAplicado < 0 THEN 0
                           ELSE cd.Saldo - a.ImporteAplicado
                       END
        FROM dbo.COM_CompraDetraccion AS cd
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'DET'
           AND a.IdRegistroComprobante = cd.IdCompraDetraccion
        WHERE cd.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE cp
        SET cp.Saldo = CASE
                           WHEN cp.Saldo - a.ImporteAplicado < 0 THEN 0
                           ELSE cp.Saldo - a.ImporteAplicado
                       END
        FROM dbo.COM_CompraPercepcion AS cp
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'PER'
           AND a.IdRegistroComprobante = cp.IdCompraPercepcion
        WHERE cp.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE cr
        SET cr.Saldo = CASE
                           WHEN cr.Saldo - a.ImporteAplicado < 0 THEN 0
                           ELSE cr.Saldo - a.ImporteAplicado
                       END
        FROM dbo.COM_CompraRetencion AS cr
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'R4T'
           AND a.IdRegistroComprobante = cr.IdCompraRetencion
        WHERE cr.IdEmpresa = @IdEmpresa;

        ;WITH AplicacionesActuales AS
        (
            SELECT
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                SUM(ISNULL(d.ImporteAplicado, 0)) AS ImporteAplicado
            FROM @Detalles AS d
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
            GROUP BY
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante
        )
        UPDATE v
        SET v.Saldo = CASE
                          WHEN v.Saldo - a.ImporteAplicado < 0 THEN 0
                          ELSE v.Saldo - a.ImporteAplicado
                      END
        FROM dbo.VEN_Venta AS v
        INNER JOIN AplicacionesActuales AS a
            ON a.ModuloOperacionComprobante = 'VEN'
           AND a.IdRegistroComprobante = v.IdVenta
        WHERE v.IdEmpresa = @IdEmpresa;

        IF @IndTranConta = 'S' AND @IdAsientoTrabajo IS NOT NULL
        BEGIN
            DECLARE @IdPlanCuentaAjuste INT
            DECLARE @ModuloOperacionAjuste CHAR(3)
            DECLARE @IdRegistroAjuste INT
            DECLARE @NumeroDocumentoAjuste VARCHAR(20)
            DECLARE @TipoDocumentoAjuste NVARCHAR(150)
            DECLARE @SerieAjuste VARCHAR(10)
            DECLARE @ReferenciaLineaAjuste NVARCHAR(100)
            DECLARE @TipoCambioLineaAjuste DECIMAL(18, 6)
            DECLARE @GlosaDetalleAjuste NVARCHAR(300)

            DECLARE cursor_ajuste_cancelacion CURSOR LOCAL FAST_FORWARD FOR
            SELECT DISTINCT
                d.IdPlanCuenta,
                d.ModuloOperacionComprobante,
                d.IdRegistroComprobante,
                d.NumeroDocumento,
                d.TipoDocumento,
                d.Serie,
                d.ReferenciaLinea,
                d.TipoCambioLinea,
                d.GlosaDetalle
            FROM @Detalles AS d
            LEFT JOIN dbo.COM_Compra AS c
                ON d.ModuloOperacionComprobante = 'COM'
               AND c.IdCompra = d.IdRegistroComprobante
               AND c.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.VEN_Venta AS v
                ON d.ModuloOperacionComprobante = 'VEN'
               AND v.IdVenta = d.IdRegistroComprobante
               AND v.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.COM_CompraDetraccion AS cd
                ON d.ModuloOperacionComprobante = 'DET'
               AND cd.IdCompraDetraccion = d.IdRegistroComprobante
               AND cd.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.COM_CompraPercepcion AS cp
                ON d.ModuloOperacionComprobante = 'PER'
               AND cp.IdCompraPercepcion = d.IdRegistroComprobante
               AND cp.IdEmpresa = @IdEmpresa
            LEFT JOIN dbo.COM_CompraRetencion AS cr
                ON d.ModuloOperacionComprobante = 'R4T'
               AND cr.IdCompraRetencion = d.IdRegistroComprobante
               AND cr.IdEmpresa = @IdEmpresa
            WHERE d.ModuloOperacionComprobante IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
              AND d.IdRegistroComprobante IS NOT NULL
              AND ISNULL(d.ImporteAplicado, 0) > 0
              AND ABS(COALESCE(c.Saldo, v.Saldo, cd.Saldo, cp.Saldo, cr.Saldo, 1)) < 0.005;

            OPEN cursor_ajuste_cancelacion;

            FETCH NEXT FROM cursor_ajuste_cancelacion
            INTO @IdPlanCuentaAjuste, @ModuloOperacionAjuste, @IdRegistroAjuste, @NumeroDocumentoAjuste, @TipoDocumentoAjuste,
                 @SerieAjuste, @ReferenciaLineaAjuste, @TipoCambioLineaAjuste, @GlosaDetalleAjuste;

            WHILE @@FETCH_STATUS = 0
            BEGIN
                EXEC dbo.usp_CON_GenerarAjusteCancelacionDiferenciaCambio
                    @IdEmpresa = @IdEmpresa,
                    @IdAsiento = @IdAsientoTrabajo,
                    @IdPlanCuentaDocumento = @IdPlanCuentaAjuste,
                    @ModuloOperacionComprobante = @ModuloOperacionAjuste,
                    @IdRegistroComprobante = @IdRegistroAjuste,
                    @NumeroDocumento = @NumeroDocumentoAjuste,
                    @TipoDocumento = @TipoDocumentoAjuste,
                    @Serie = @SerieAjuste,
                    @ReferenciaLinea = @ReferenciaLineaAjuste,
                    @TipoCambioLinea = @TipoCambioLineaAjuste,
                    @GlosaDetalle = @GlosaDetalleAjuste,
                    @UsuarioRegistro = @UsuarioRegistro;

                FETCH NEXT FROM cursor_ajuste_cancelacion
                INTO @IdPlanCuentaAjuste, @ModuloOperacionAjuste, @IdRegistroAjuste, @NumeroDocumentoAjuste, @TipoDocumentoAjuste,
                     @SerieAjuste, @ReferenciaLineaAjuste, @TipoCambioLineaAjuste, @GlosaDetalleAjuste;
            END;

            CLOSE cursor_ajuste_cancelacion;
            DEALLOCATE cursor_ajuste_cancelacion;
        END;

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

        IF CURSOR_STATUS('local', 'cursor_ajuste_cancelacion') >= -1
        BEGIN
            CLOSE cursor_ajuste_cancelacion;
            DEALLOCATE cursor_ajuste_cancelacion;
        END;

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
