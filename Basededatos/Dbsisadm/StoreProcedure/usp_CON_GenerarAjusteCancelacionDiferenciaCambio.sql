-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/07/2026
-- Description:   Genera lineas analiticas de ajuste por cancelacion total de documentos para saldar diferencias residuales en soles y dolares sin alterar el Debe/Haber del asiento.
-- =============================================
-- Firma: FRANCO LARA - 06/07/2026 | Crea el ajuste cambiario por cancelacion total desde asientos operativos, agregando lineas separadas por soles y dolares con Debe/Haber en cero, DH para identificar ganancia o perdida y expansion final de cuentas destino/contrapartida cuando la cuenta ajustada las tenga configuradas.
-- Firma: FRANCO LARA - 11/07/2026 | Agrega la contrapartida analitica sobre la cuenta del comprobante para que cada ganancia o perdida por cancelacion total quede compensada tambien en la 42/12 correspondiente usando solo TotalImporteS y/o TotalImporteD con Debe/Haber en cero.
-- Firma: FRANCO LARA - 13/07/2026 | Corrige el ajuste por cancelacion total para que la linea del documento invierta el DH del residual, y la linea de ganancia o perdida use el signo correcto segun el saldo analitico remanente en soles o dolares.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GenerarAjusteCancelacionDiferenciaCambio
    @IdEmpresa INT,
    @IdAsiento INT,
    @IdPlanCuentaDocumento INT,
    @ModuloOperacionComprobante CHAR(3),
    @IdRegistroComprobante INT,
    @NumeroDocumento VARCHAR(20) = NULL,
    @TipoDocumento NVARCHAR(150) = NULL,
    @Serie VARCHAR(10) = NULL,
    @ReferenciaLinea NVARCHAR(100) = NULL,
    @TipoCambioLinea DECIMAL(18, 6) = NULL,
    @GlosaDetalle NVARCHAR(300) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @CodigoCuentaGanancia VARCHAR(20)
        DECLARE @CodigoCuentaPerdida VARCHAR(20)
        DECLARE @IdPlanCuentaGanancia INT
        DECLARE @IdPlanCuentaPerdida INT
        DECLARE @SaldoDocumento DECIMAL(18, 2) = NULL
        DECLARE @TipoCambioLineaTrabajo DECIMAL(18, 6) = NULL
        DECLARE @TotalSoles DECIMAL(18, 2) = 0
        DECLARE @TotalDolares DECIMAL(18, 2) = 0
        DECLARE @ImporteSolesAjuste DECIMAL(18, 2) = 0
        DECLARE @ImporteDolaresAjuste DECIMAL(18, 2) = 0
        DECLARE @DhSoles CHAR(1) = NULL
        DECLARE @DhDolares CHAR(1) = NULL
        DECLARE @SiguienteItem INT
        DECLARE @LineasAjusteGeneradas TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            DH CHAR(1) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL,
            NumeroDocumento VARCHAR(20) NULL,
            TipoDocumento NVARCHAR(150) NULL,
            Serie VARCHAR(10) NULL,
            ReferenciaLinea NVARCHAR(100) NULL,
            TipoCambioLinea DECIMAL(18, 6) NOT NULL,
            TotalImporteS DECIMAL(18, 2) NOT NULL,
            TotalImporteD DECIMAL(18, 2) NOT NULL
        )
        DECLARE @CuentaDestinoDetalle TABLE
        (
            IdPlanCuentaOrigen INT NOT NULL,
            Orden INT NOT NULL,
            IdPlanCuentaDestinoCargo INT NOT NULL,
            IdPlanCuentaDestinoAbono INT NOT NULL,
            Porcentaje DECIMAL(7, 4) NOT NULL,
            EsUltimo BIT NOT NULL
        )

        IF ISNULL(@IdAsiento, 0) <= 0
           OR ISNULL(@IdPlanCuentaDocumento, 0) <= 0
           OR ISNULL(@IdRegistroComprobante, 0) <= 0
           OR @ModuloOperacionComprobante NOT IN ('COM', 'VEN', 'DET', 'PER', 'R4T')
        BEGIN
            RETURN;
        END;

        SELECT
            @TipoCambioLineaTrabajo = CASE
                                          WHEN ISNULL(@TipoCambioLinea, 0) > 0 THEN @TipoCambioLinea
                                          ELSE a.TipoCambio
                                      END
        FROM dbo.CON_Asiento AS a
        WHERE a.IdAsiento = @IdAsiento
          AND a.IdEmpresa = @IdEmpresa;

        IF ISNULL(@TipoCambioLineaTrabajo, 0) <= 0
        BEGIN
            SET @TipoCambioLineaTrabajo = 1;
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
            @SaldoDocumento = CASE @ModuloOperacionComprobante
                                  WHEN 'COM' THEN c.Saldo
                                  WHEN 'VEN' THEN v.Saldo
                                  WHEN 'DET' THEN cd.Saldo
                                  WHEN 'PER' THEN cp.Saldo
                                  WHEN 'R4T' THEN cr.Saldo
                              END
        FROM (SELECT 1 AS Dummy) AS x
        LEFT JOIN dbo.COM_Compra AS c
            ON @ModuloOperacionComprobante = 'COM'
           AND c.IdCompra = @IdRegistroComprobante
           AND c.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.VEN_Venta AS v
            ON @ModuloOperacionComprobante = 'VEN'
           AND v.IdVenta = @IdRegistroComprobante
           AND v.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.COM_CompraDetraccion AS cd
            ON @ModuloOperacionComprobante = 'DET'
           AND cd.IdCompraDetraccion = @IdRegistroComprobante
           AND cd.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.COM_CompraPercepcion AS cp
            ON @ModuloOperacionComprobante = 'PER'
           AND cp.IdCompraPercepcion = @IdRegistroComprobante
           AND cp.IdEmpresa = @IdEmpresa
        LEFT JOIN dbo.COM_CompraRetencion AS cr
            ON @ModuloOperacionComprobante = 'R4T'
           AND cr.IdCompraRetencion = @IdRegistroComprobante
           AND cr.IdEmpresa = @IdEmpresa;

        IF ABS(ISNULL(@SaldoDocumento, 0)) >= 0.005
        BEGIN
            RETURN;
        END;

        SELECT
            @TotalSoles = ROUND(ISNULL(SUM(CASE WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteS, 0) ELSE ISNULL(d.TotalImporteS, 0) * -1 END), 0), 2),
            @TotalDolares = ROUND(ISNULL(SUM(CASE WHEN d.DH = 'D' THEN ISNULL(d.TotalImporteD, 0) ELSE ISNULL(d.TotalImporteD, 0) * -1 END), 0), 2)
        FROM dbo.CON_AsientoDetalle AS d
        INNER JOIN dbo.CON_Asiento AS a
            ON a.IdAsiento = d.IdAsiento
        WHERE a.IdEmpresa = @IdEmpresa
          AND (
                d.IdPlanCuenta = @IdPlanCuentaDocumento
                OR d.IdPlanCuenta = @IdPlanCuentaGanancia
                OR d.IdPlanCuenta = @IdPlanCuentaPerdida
              )
          AND ISNULL(NULLIF(LTRIM(RTRIM(d.NumeroDocumento)), ''), '') = ISNULL(NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''), '')
          AND ISNULL(NULLIF(LTRIM(RTRIM(d.TipoDocumento)), N''), N'') = ISNULL(NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''), N'')
          AND ISNULL(NULLIF(LTRIM(RTRIM(d.Serie)), ''), '') = ISNULL(NULLIF(LTRIM(RTRIM(@Serie)), ''), '')
          AND ISNULL(NULLIF(LTRIM(RTRIM(d.ReferenciaLinea)), N''), N'') = ISNULL(NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''), N'');

        IF ABS(@TotalSoles) < 0.005 AND ABS(@TotalDolares) < 0.005
        BEGIN
            RETURN;
        END;

        SET @ImporteSolesAjuste = CASE WHEN ABS(@TotalSoles) >= 0.005 THEN ABS(@TotalSoles) ELSE 0 END;
        SET @ImporteDolaresAjuste = CASE WHEN ABS(@TotalDolares) >= 0.005 THEN ABS(@TotalDolares) ELSE 0 END;
        SET @DhSoles = CASE WHEN @TotalSoles > 0 THEN 'D' WHEN @TotalSoles < 0 THEN 'H' ELSE NULL END;
        SET @DhDolares = CASE WHEN @TotalDolares > 0 THEN 'D' WHEN @TotalDolares < 0 THEN 'H' ELSE NULL END;

        SELECT
            @SiguienteItem = ISNULL(MAX(d.Item), 0) + 1
        FROM dbo.CON_AsientoDetalle AS d
        WHERE d.IdAsiento = @IdAsiento;

        IF @ImporteSolesAjuste > 0 AND @DhSoles IS NOT NULL
        BEGIN
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
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD,
                UsuarioRegistro
            )
            VALUES
            (
                @IdAsiento,
                @SiguienteItem,
                @IdPlanCuentaDocumento,
                CASE WHEN @DhSoles = 'D' THEN 'H' ELSE 'D' END,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE SOLES DOC'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                0,
                0,
                @ImporteSolesAjuste,
                0,
                @UsuarioRegistro
            );

            SET @SiguienteItem += 1;

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
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD,
                UsuarioRegistro
            )
            VALUES
            (
                @IdAsiento,
                @SiguienteItem,
                CASE WHEN @DhSoles = 'D' THEN @IdPlanCuentaPerdida ELSE @IdPlanCuentaGanancia END,
                @DhSoles,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE SOLES'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                0,
                0,
                @ImporteSolesAjuste,
                0,
                @UsuarioRegistro
            );

            INSERT INTO @LineasAjusteGeneradas
            (
                IdPlanCuentaOrigen,
                DH,
                GlosaDetalle,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                TotalImporteS,
                TotalImporteD
            )
            VALUES
            (
                CASE WHEN @DhSoles = 'D' THEN @IdPlanCuentaPerdida ELSE @IdPlanCuentaGanancia END,
                @DhSoles,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE SOLES'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                @ImporteSolesAjuste,
                0
            );

            SET @SiguienteItem += 1;
        END;

        IF @ImporteDolaresAjuste > 0 AND @DhDolares IS NOT NULL
        BEGIN
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
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD,
                UsuarioRegistro
            )
            VALUES
            (
                @IdAsiento,
                @SiguienteItem,
                @IdPlanCuentaDocumento,
                CASE WHEN @DhDolares = 'D' THEN 'H' ELSE 'D' END,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE DOLARES DOC'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                0,
                0,
                0,
                @ImporteDolaresAjuste,
                @UsuarioRegistro
            );

            SET @SiguienteItem += 1;

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
                Debe,
                Haber,
                TotalImporteS,
                TotalImporteD,
                UsuarioRegistro
            )
            VALUES
            (
                @IdAsiento,
                @SiguienteItem,
                CASE WHEN @DhDolares = 'D' THEN @IdPlanCuentaPerdida ELSE @IdPlanCuentaGanancia END,
                @DhDolares,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE DOLARES'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                0,
                0,
                0,
                @ImporteDolaresAjuste,
                @UsuarioRegistro
            );

            INSERT INTO @LineasAjusteGeneradas
            (
                IdPlanCuentaOrigen,
                DH,
                GlosaDetalle,
                NumeroDocumento,
                TipoDocumento,
                Serie,
                ReferenciaLinea,
                TipoCambioLinea,
                TotalImporteS,
                TotalImporteD
            )
            VALUES
            (
                CASE WHEN @DhDolares = 'D' THEN @IdPlanCuentaPerdida ELSE @IdPlanCuentaGanancia END,
                @DhDolares,
                LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaDetalle)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / AJUSTE DOLARES'), 300),
                NULLIF(LTRIM(RTRIM(@NumeroDocumento)), ''),
                NULLIF(LTRIM(RTRIM(@TipoDocumento)), N''),
                NULLIF(LTRIM(RTRIM(@Serie)), ''),
                NULLIF(LTRIM(RTRIM(@ReferenciaLinea)), N''),
                @TipoCambioLineaTrabajo,
                0,
                @ImporteDolaresAjuste
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
                l.IdPlanCuentaOrigen
            FROM @LineasAjusteGeneradas AS l
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
            RAISERROR(N'Existe una configuracion activa de cuentas destino con cuentas cargo o abono invalidas para la empresa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @CuentaDestinoDetalle
        )
        BEGIN
            DECLARE @IdPlanCuentaOrigenDestino INT
            DECLARE @DhOrigenDestino CHAR(1)
            DECLARE @GlosaBaseDestino NVARCHAR(300)
            DECLARE @NumeroDocumentoDestino VARCHAR(20)
            DECLARE @TipoDocumentoDestino NVARCHAR(150)
            DECLARE @SerieDestino VARCHAR(10)
            DECLARE @ReferenciaLineaDestino NVARCHAR(100)
            DECLARE @TipoCambioLineaDestino DECIMAL(18, 6)
            DECLARE @TotalImporteSOrigenDestino DECIMAL(18, 2)
            DECLARE @TotalImporteDOrigenDestino DECIMAL(18, 2)
            DECLARE @IdCuentaCargoDestino INT
            DECLARE @IdCuentaAbonoDestino INT
            DECLARE @PorcentajeDestino DECIMAL(7, 4)
            DECLARE @EsUltimoDestino BIT
            DECLARE @ImporteBaseSolesDestino DECIMAL(18, 2)
            DECLARE @ImporteBaseDolaresDestino DECIMAL(18, 2)
            DECLARE @ImporteDistribuidoSolesDestino DECIMAL(18, 2)
            DECLARE @ImporteDistribuidoDolaresDestino DECIMAL(18, 2)
            DECLARE @ImporteTramoSolesDestino DECIMAL(18, 2)
            DECLARE @ImporteTramoDolaresDestino DECIMAL(18, 2)

            DECLARE cursor_linea_destino CURSOR LOCAL FAST_FORWARD FOR
            SELECT
                l.IdPlanCuentaOrigen,
                l.DH,
                l.GlosaDetalle,
                l.NumeroDocumento,
                l.TipoDocumento,
                l.Serie,
                l.ReferenciaLinea,
                l.TipoCambioLinea,
                l.TotalImporteS,
                l.TotalImporteD
            FROM @LineasAjusteGeneradas AS l
            WHERE EXISTS
            (
                SELECT 1
                FROM @CuentaDestinoDetalle AS r
                WHERE r.IdPlanCuentaOrigen = l.IdPlanCuentaOrigen
            );

            OPEN cursor_linea_destino;

            FETCH NEXT FROM cursor_linea_destino
            INTO @IdPlanCuentaOrigenDestino, @DhOrigenDestino, @GlosaBaseDestino, @NumeroDocumentoDestino, @TipoDocumentoDestino, @SerieDestino,
                 @ReferenciaLineaDestino, @TipoCambioLineaDestino, @TotalImporteSOrigenDestino, @TotalImporteDOrigenDestino;

            WHILE @@FETCH_STATUS = 0
            BEGIN
                SET @ImporteBaseSolesDestino = ISNULL(@TotalImporteSOrigenDestino, 0);
                SET @ImporteBaseDolaresDestino = ISNULL(@TotalImporteDOrigenDestino, 0);
                SET @ImporteDistribuidoSolesDestino = 0;
                SET @ImporteDistribuidoDolaresDestino = 0;

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
                    SET @ImporteTramoSolesDestino = CASE
                                                        WHEN @ImporteBaseSolesDestino = 0 THEN 0
                                                        WHEN @EsUltimoDestino = 1 THEN @ImporteBaseSolesDestino - @ImporteDistribuidoSolesDestino
                                                        ELSE ROUND(@ImporteBaseSolesDestino * (@PorcentajeDestino / 100.0), 2)
                                                    END;

                    SET @ImporteTramoDolaresDestino = CASE
                                                          WHEN @ImporteBaseDolaresDestino = 0 THEN 0
                                                          WHEN @EsUltimoDestino = 1 THEN @ImporteBaseDolaresDestino - @ImporteDistribuidoDolaresDestino
                                                          ELSE ROUND(@ImporteBaseDolaresDestino * (@PorcentajeDestino / 100.0), 2)
                                                      END;

                    IF @ImporteTramoSolesDestino <> 0 OR @ImporteTramoDolaresDestino <> 0
                    BEGIN
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
                            Debe,
                            Haber,
                            TotalImporteS,
                            TotalImporteD,
                            UsuarioRegistro
                        )
                        VALUES
                        (
                            @IdAsiento,
                            @SiguienteItem,
                            @IdCuentaCargoDestino,
                            'D',
                            LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / DESTINO'), 300),
                            @NumeroDocumentoDestino,
                            @TipoDocumentoDestino,
                            @SerieDestino,
                            @ReferenciaLineaDestino,
                            @TipoCambioLineaDestino,
                            0,
                            0,
                            @ImporteTramoSolesDestino,
                            @ImporteTramoDolaresDestino,
                            @UsuarioRegistro
                        );

                        SET @SiguienteItem += 1;

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
                            Debe,
                            Haber,
                            TotalImporteS,
                            TotalImporteD,
                            UsuarioRegistro
                        )
                        VALUES
                        (
                            @IdAsiento,
                            @SiguienteItem,
                            @IdCuentaAbonoDestino,
                            'H',
                            LEFT(CONCAT(ISNULL(NULLIF(LTRIM(RTRIM(@GlosaBaseDestino)), N''), N'AJUSTE CANCELACION DIF. CAMBIO'), N' / CONTRAPARTIDA'), 300),
                            @NumeroDocumentoDestino,
                            @TipoDocumentoDestino,
                            @SerieDestino,
                            @ReferenciaLineaDestino,
                            @TipoCambioLineaDestino,
                            0,
                            0,
                            @ImporteTramoSolesDestino,
                            @ImporteTramoDolaresDestino,
                            @UsuarioRegistro
                        );

                        SET @SiguienteItem += 1;
                    END;

                    SET @ImporteDistribuidoSolesDestino = @ImporteDistribuidoSolesDestino + @ImporteTramoSolesDestino;
                    SET @ImporteDistribuidoDolaresDestino = @ImporteDistribuidoDolaresDestino + @ImporteTramoDolaresDestino;

                    FETCH NEXT FROM cursor_tramo_destino
                    INTO @IdCuentaCargoDestino, @IdCuentaAbonoDestino, @PorcentajeDestino, @EsUltimoDestino;
                END;

                CLOSE cursor_tramo_destino;
                DEALLOCATE cursor_tramo_destino;

                FETCH NEXT FROM cursor_linea_destino
                INTO @IdPlanCuentaOrigenDestino, @DhOrigenDestino, @GlosaBaseDestino, @NumeroDocumentoDestino, @TipoDocumentoDestino, @SerieDestino,
                     @ReferenciaLineaDestino, @TipoCambioLineaDestino, @TotalImporteSOrigenDestino, @TotalImporteDOrigenDestino;
            END;

            CLOSE cursor_linea_destino;
            DEALLOCATE cursor_linea_destino;
        END;

    END TRY

    BEGIN CATCH

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
