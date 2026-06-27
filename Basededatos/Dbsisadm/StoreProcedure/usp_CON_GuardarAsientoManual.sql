-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza un asiento manual por empresa validando cuadre, periodo y correlativo mensual.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Incluye centro de costo, documento, serie y TC opcional en el detalle del asiento.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Aclara la validacion de cuentas, amplia TipoDocumento para soportar descripciones de comprobante en asientos editados y agrega fecha de emision permitiendo asientos descuadrados.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Valida centros de costo activos por empresa y exige su registro cuando la cuenta contable lo requiere.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/06/2026
-- Description:   Conserva la linea original del detalle, agrega cuentas destino y contrapartida segun la configuracion activa y deja el estado del asiento como PROVISIONADO.
-- =============================================
-- Firma: FRANCO LARA - 26/06/2026 | Guarda tipo documento por codigo y calcula equivalencias en soles y dolares por cada linea del asiento manual.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarAsientoManual
    @IdAsiento INT = NULL,
    @IdEmpresa INT,
    @IdOrigen INT,
    @FechaEmision DATE,
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
        DECLARE @CodigoMoneda VARCHAR(10)
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
            CodigoCentroCosto NVARCHAR(50) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            Serie VARCHAR(10) NULL,
            TipoCambioLinea DECIMAL(18,6) NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            ReferenciaLinea NVARCHAR(100) NULL
        );

        DECLARE @DetalleExpandido TABLE
        (
            Item INT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            CodigoCentroCosto NVARCHAR(50) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            Serie VARCHAR(10) NULL,
            TipoCambioLinea DECIMAL(18,6) NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            ReferenciaLinea NVARCHAR(100) NULL
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

        INSERT INTO @Detalle
        (
            Item,
            IdPlanCuenta,
            GlosaDetalle,
            CodigoCentroCosto,
            TipoDocumento,
            NumeroDocumento,
            Serie,
            TipoCambioLinea,
            Debe,
            Haber,
            ReferenciaLinea
        )
        SELECT
            T.N.value('@Item', 'smallint'),
            T.N.value('@IdPlanCuenta', 'int'),
            NULLIF(T.N.value('@GlosaDetalle', 'nvarchar(300)'), N''),
            NULLIF(LTRIM(RTRIM(T.N.value('@CodigoCentroCosto', 'nvarchar(50)'))), N''),
            NULLIF(T.N.value('@TipoDocumento', 'nvarchar(150)'), N''),
            NULLIF(T.N.value('@NumeroDocumento', 'varchar(20)'), ''),
            NULLIF(T.N.value('@Serie', 'varchar(10)'), ''),
            NULLIF(T.N.value('@TipoCambioLinea', 'decimal(18,6)'), 0),
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
            RAISERROR(N'Todas las cuentas del detalle deben pertenecer a la empresa, estar activas y aceptar movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            INNER JOIN dbo.CON_PlanCuenta AS p
                ON p.IdPlanCuenta = d.IdPlanCuenta
               AND p.IdEmpresa = @IdEmpresa
            WHERE p.RequiereCentroCosto = 1
              AND NULLIF(LTRIM(RTRIM(d.CodigoCentroCosto)), N'') IS NULL
        )
        BEGIN
            RAISERROR(N'Las cuentas configuradas con centro de costo obligatorio deben registrar un centro de costo.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            LEFT JOIN dbo.CON_CentroCostoConfiguracionEmpresa AS c
                ON c.IdEmpresa = @IdEmpresa
               AND c.Codigo = d.CodigoCentroCosto
               AND c.Estado = 1
            WHERE NULLIF(LTRIM(RTRIM(d.CodigoCentroCosto)), N'') IS NOT NULL
              AND c.IdCentroCostoConfiguracionEmpresa IS NULL
        )
        BEGIN
            RAISERROR(N'Todo centro de costo informado debe existir y estar activo para la empresa.', 16, 1);
        END;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @Detalle AS d;

        IF ISNULL(@TotalDebe, 0) <= 0 AND ISNULL(@TotalHaber, 0) <= 0
        BEGIN
            RAISERROR(N'El asiento debe tener al menos un importe positivo en el detalle.', 16, 1);
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

        SELECT
            @CodigoMoneda = m.CodigoMoneda
        FROM dbo.ADM_Moneda AS m
        WHERE m.IdMoneda = @IdMoneda;

        INSERT INTO @DetalleExpandido
        (
            IdPlanCuenta,
            GlosaDetalle,
            CodigoCentroCosto,
            TipoDocumento,
            NumeroDocumento,
            Serie,
            TipoCambioLinea,
            Debe,
            Haber,
            ReferenciaLinea
        )
        SELECT
            d.IdPlanCuenta,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.TipoDocumento,
            d.NumeroDocumento,
            d.Serie,
            d.TipoCambioLinea,
            d.Debe,
            d.Haber,
            d.ReferenciaLinea
        FROM @Detalle AS d
        ORDER BY
            d.Item ASC;

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
                x.IdPlanCuenta
            FROM @Detalle AS x
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

        DECLARE @IdPlanCuentaOrigenDestino INT
        DECLARE @GlosaBaseDestino NVARCHAR(300)
        DECLARE @CodigoCentroCostoDestino NVARCHAR(50)
        DECLARE @TipoDocumentoDestino NVARCHAR(150)
        DECLARE @NumeroDocumentoDestino VARCHAR(20)
        DECLARE @SerieDestino VARCHAR(10)
        DECLARE @TipoCambioLineaDestino DECIMAL(18,6)
        DECLARE @DebeOrigenDestino DECIMAL(18,2)
        DECLARE @HaberOrigenDestino DECIMAL(18,2)
        DECLARE @ReferenciaLineaDestino NVARCHAR(100)
        DECLARE @ImporteBaseDestino DECIMAL(18,2)
        DECLARE @IdCuentaCargoDestino INT
        DECLARE @IdCuentaAbonoDestino INT
        DECLARE @PorcentajeDestino DECIMAL(7,4)
        DECLARE @EsUltimoDestino BIT
        DECLARE @ImporteDistribuidoDestino DECIMAL(18,2)
        DECLARE @ImporteTramoDestino DECIMAL(18,2)

        DECLARE cursor_linea_destino CURSOR LOCAL FAST_FORWARD FOR
        SELECT
            d.IdPlanCuenta,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.TipoDocumento,
            d.NumeroDocumento,
            d.Serie,
            d.TipoCambioLinea,
            d.Debe,
            d.Haber,
            d.ReferenciaLinea
        FROM @Detalle AS d
        WHERE (d.Debe > 0 OR d.Haber > 0)
          AND EXISTS
          (
              SELECT 1
              FROM @CuentaDestinoDetalle AS r
              WHERE r.IdPlanCuentaOrigen = d.IdPlanCuenta
          )
        ORDER BY
            d.Item ASC;

        OPEN cursor_linea_destino;

        FETCH NEXT FROM cursor_linea_destino
        INTO @IdPlanCuentaOrigenDestino, @GlosaBaseDestino, @CodigoCentroCostoDestino, @TipoDocumentoDestino,
             @NumeroDocumentoDestino, @SerieDestino, @TipoCambioLineaDestino, @DebeOrigenDestino, @HaberOrigenDestino, @ReferenciaLineaDestino;

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
                    INSERT INTO @DetalleExpandido
                    (
                        IdPlanCuenta,
                        GlosaDetalle,
                        CodigoCentroCosto,
                        TipoDocumento,
                        NumeroDocumento,
                        Serie,
                        TipoCambioLinea,
                        Debe,
                        Haber,
                        ReferenciaLinea
                    )
                    VALUES
                    (
                        @IdCuentaCargoDestino,
                        LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'Distribucion cuenta destino'), N' / Destino'), 300),
                        @CodigoCentroCostoDestino,
                        @TipoDocumentoDestino,
                        @NumeroDocumentoDestino,
                        @SerieDestino,
                        @TipoCambioLineaDestino,
                        @ImporteTramoDestino,
                        0,
                        @ReferenciaLineaDestino
                    );

                    INSERT INTO @DetalleExpandido
                    (
                        IdPlanCuenta,
                        GlosaDetalle,
                        CodigoCentroCosto,
                        TipoDocumento,
                        NumeroDocumento,
                        Serie,
                        TipoCambioLinea,
                        Debe,
                        Haber,
                        ReferenciaLinea
                    )
                    VALUES
                    (
                        @IdCuentaAbonoDestino,
                        LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'Distribucion cuenta destino'), N' / Contrapartida'), 300),
                        @CodigoCentroCostoDestino,
                        @TipoDocumentoDestino,
                        @NumeroDocumentoDestino,
                        @SerieDestino,
                        @TipoCambioLineaDestino,
                        0,
                        @ImporteTramoDestino,
                        @ReferenciaLineaDestino
                    );
                END;

                SET @ImporteDistribuidoDestino = @ImporteDistribuidoDestino + @ImporteTramoDestino;

                FETCH NEXT FROM cursor_tramo_destino
                INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
            END;

            CLOSE cursor_tramo_destino;
            DEALLOCATE cursor_tramo_destino;

            FETCH NEXT FROM cursor_linea_destino
            INTO @IdPlanCuentaOrigenDestino, @GlosaBaseDestino, @CodigoCentroCostoDestino, @TipoDocumentoDestino,
                 @NumeroDocumentoDestino, @SerieDestino, @TipoCambioLineaDestino, @DebeOrigenDestino, @HaberOrigenDestino, @ReferenciaLineaDestino;
        END;

        CLOSE cursor_linea_destino;
        DEALLOCATE cursor_linea_destino;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @DetalleExpandido AS d;

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
                @FechaAsiento,
                @Glosa,
                @IdMoneda,
                @TipoCambio,
                @TotalDebe,
                @TotalHaber,
                N'PROVISIONADO',
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
            SET FechaEmision = @FechaEmision,
                FechaAsiento = @FechaAsiento,
                Glosa = @Glosa,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebe,
                TotalHaber = @TotalHaber,
                Estado = N'PROVISIONADO',
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
            CodigoCentroCosto,
            TipoDocumento,
            NumeroDocumento,
            Serie,
            TipoCambioLinea,
            Debe,
            Haber,
            TotalImporteS,
            TotalImporteD,
            ReferenciaLinea,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.GlosaDetalle,
            d.CodigoCentroCosto,
            d.TipoDocumento,
            d.NumeroDocumento,
            d.Serie,
            calc.TipoCambioAplicado,
            d.Debe,
            d.Haber,
            CASE
                WHEN @CodigoMoneda = 'USD' THEN ROUND(calc.ImporteLinea * calc.TipoCambioAplicado, 2)
                ELSE calc.ImporteLinea
            END,
            CASE
                WHEN @CodigoMoneda = 'USD' THEN calc.ImporteLinea
                ELSE ROUND(calc.ImporteLinea / NULLIF(calc.TipoCambioAplicado, 0), 2)
            END,
            d.ReferenciaLinea,
            @UsuarioRegistro
        FROM @DetalleExpandido AS d
        CROSS APPLY
        (
            SELECT
                CASE
                    WHEN d.Debe > 0 THEN d.Debe
                    ELSE d.Haber
                END AS ImporteLinea,
                ISNULL(NULLIF(d.TipoCambioLinea, 0), CASE WHEN @TipoCambio > 0 THEN @TipoCambio ELSE 1 END) AS TipoCambioAplicado
        ) AS calc
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
