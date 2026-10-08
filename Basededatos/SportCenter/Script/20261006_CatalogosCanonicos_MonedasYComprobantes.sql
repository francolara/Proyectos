
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Introduce codigos canonicos de moneda y documento en operaciones historicas.
-- Firma:         FRANCO LARA - 06/10/2026 | Agrega catalogos canonicos, migra historicos y deja la configuracion preparada para retirar IDs locales.
-- =============================================

SET XACT_ABORT ON;
BEGIN TRANSACTION;

BEGIN TRY
    IF EXISTS
    (
        SELECT 1
        FROM dbo.MonedasSuperMaestro msm
        GROUP BY msm.Codigo
        HAVING COUNT(*) > 1
    )
        RAISERROR('No se puede crear el catalogo canonico: existen codigos duplicados en MonedasSuperMaestro.', 16, 1);

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.indexes
        WHERE object_id = OBJECT_ID(N'dbo.MonedasSuperMaestro')
          AND name = N'UQ_MonedasSuperMaestro_Codigo'
    )
        CREATE UNIQUE NONCLUSTERED INDEX UQ_MonedasSuperMaestro_Codigo
            ON dbo.MonedasSuperMaestro (Codigo);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Monedas m
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = m.Codigo
        WHERE msm.Codigo IS NULL
    )
        RAISERROR('Existen monedas configuradas por negocio sin un Codigo valido en MonedasSuperMaestro.', 16, 1);

    IF COL_LENGTH(N'dbo.Negocios', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.Negocios ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.Reservas', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.Reservas ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.Pagos', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.Pagos ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.Tarifas', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.Tarifas ADD CodigoMoneda NVARCHAR(10) NULL;

    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
       AND COL_LENGTH(N'dbo.TarifaFeriado', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.TarifaFeriado ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.Cupones', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.Cupones ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.CuponesUso', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.CuponesUso ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.ComprobantesElectronicos ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoTipoComprobante') IS NULL
        ALTER TABLE dbo.ComprobantesElectronicos ADD CodigoTipoComprobante NVARCHAR(4) NULL;

    IF COL_LENGTH(N'dbo.ComprobantesDetalle', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.ComprobantesDetalle ADD CodigoMoneda NVARCHAR(10) NULL;

    IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'CodigoMoneda') IS NULL
        ALTER TABLE dbo.NegociosSuscripcionPago ADD CodigoMoneda NVARCHAR(10) NULL;

    UPDATE n
       SET CodigoMoneda = COALESCE(msm.Codigo, m.Codigo)
    FROM dbo.Negocios n
    LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
    LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Id = m.MonedaSuperId
    WHERE n.CodigoMoneda IS NULL;

    UPDATE r
       SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Reservas r
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
    WHERE r.CodigoMoneda IS NULL;

    UPDATE p
       SET CodigoMoneda = r.CodigoMoneda
    FROM dbo.Pagos p
    INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
    WHERE p.CodigoMoneda IS NULL;

    UPDATE t
       SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Tarifas t
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = t.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
    WHERE t.CodigoMoneda IS NULL;

    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
        UPDATE tf
           SET CodigoMoneda = n.CodigoMoneda
        FROM dbo.TarifaFeriado tf
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = tf.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE tf.CodigoMoneda IS NULL;

    UPDATE c
       SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Cupones c
    INNER JOIN dbo.Negocios n ON n.Id = c.NegocioId
    WHERE c.CodigoMoneda IS NULL;

    UPDATE cu
       SET CodigoMoneda = r.CodigoMoneda
    FROM dbo.CuponesUso cu
    INNER JOIN dbo.Reservas r ON r.Id = cu.ReservaId
    WHERE cu.CodigoMoneda IS NULL;

    UPDATE ce
       SET CodigoMoneda = COALESCE(msm.Codigo, m.Codigo),
           CodigoTipoComprobante = ntd.CodigoSunat
    FROM dbo.ComprobantesElectronicos ce
    LEFT JOIN dbo.Monedas m ON m.Id = ce.TipoMoneda
    LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Id = m.MonedaSuperId
    LEFT JOIN dbo.NegociosTiposDocumentoComprobante ntd
        ON ntd.Id = ce.TipoComprobante
       AND ntd.NegocioId = ce.NegocioId
    WHERE ce.CodigoMoneda IS NULL
       OR ce.CodigoTipoComprobante IS NULL;

    UPDATE cd
       SET CodigoMoneda = ce.CodigoMoneda
    FROM dbo.ComprobantesDetalle cd
    INNER JOIN dbo.ComprobantesElectronicos ce ON ce.Id = cd.ComprobanteElectronicoId
    WHERE cd.CodigoMoneda IS NULL;

    UPDATE nsp
       SET CodigoMoneda = UPPER(LTRIM(RTRIM(nsp.Moneda)))
    FROM dbo.NegociosSuscripcionPago nsp
    WHERE nsp.CodigoMoneda IS NULL;

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Negocios n
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = n.CodigoMoneda
        WHERE n.CodigoMoneda IS NOT NULL
          AND msm.Codigo IS NULL
    )
        RAISERROR('Existen negocios con CodigoMoneda canonico invalido. Corrija la configuracion de moneda antes de continuar.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Reservas r
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = r.CodigoMoneda
        WHERE r.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen reservas sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Pagos p
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = p.CodigoMoneda
        WHERE p.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen pagos sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Tarifas t
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = t.CodigoMoneda
        WHERE t.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen tarifas sin CodigoMoneda canonico valido.', 16, 1);

    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
       AND EXISTS
       (
            SELECT 1
            FROM dbo.TarifaFeriado tf
            LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = tf.CodigoMoneda
            WHERE tf.CodigoMoneda IS NULL OR msm.Codigo IS NULL
       )
        RAISERROR('Existen tarifas de feriado sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Cupones c
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = c.CodigoMoneda
        WHERE c.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen cupones sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.CuponesUso cu
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = cu.CodigoMoneda
        WHERE cu.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen usos de cupon sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.ComprobantesElectronicos ce
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = ce.CodigoMoneda
        LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro tdsm ON tdsm.CodigoSunat = ce.CodigoTipoComprobante
        WHERE ce.CodigoMoneda IS NULL
           OR msm.Codigo IS NULL
           OR ce.CodigoTipoComprobante IS NULL
           OR tdsm.CodigoSunat IS NULL
    )
        RAISERROR('Existen comprobantes sin codigo canonico de moneda o documento.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.ComprobantesDetalle cd
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = cd.CodigoMoneda
        WHERE cd.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen detalles de comprobante sin CodigoMoneda canonico valido.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.NegociosSuscripcionPago nsp
        LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = nsp.CodigoMoneda
        WHERE nsp.CodigoMoneda IS NULL OR msm.Codigo IS NULL
    )
        RAISERROR('Existen cobros de suscripcion sin CodigoMoneda canonico valido.', 16, 1);

    -- Fase 1: las nuevas columnas permanecen nullable para compatibilidad con
    -- procedimientos aun no migrados. La Fase 2 las volvera obligatorias solo
    -- despues de actualizar todas las rutas de escritura transaccional.

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Negocios_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.Negocios WITH CHECK ADD CONSTRAINT FK_Negocios_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Monedas_MonedasSuperMaestro_Codigo')
        ALTER TABLE dbo.Monedas WITH CHECK ADD CONSTRAINT FK_Monedas_MonedasSuperMaestro_Codigo
            FOREIGN KEY (Codigo) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Reservas_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.Reservas WITH CHECK ADD CONSTRAINT FK_Reservas_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Pagos_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.Pagos WITH CHECK ADD CONSTRAINT FK_Pagos_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Tarifas_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.Tarifas WITH CHECK ADD CONSTRAINT FK_Tarifas_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
       AND NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_TarifaFeriado_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.TarifaFeriado WITH CHECK ADD CONSTRAINT FK_TarifaFeriado_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_Cupones_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.Cupones WITH CHECK ADD CONSTRAINT FK_Cupones_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_CuponesUso_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.CuponesUso WITH CHECK ADD CONSTRAINT FK_CuponesUso_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ComprobantesElectronicos_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.ComprobantesElectronicos WITH CHECK ADD CONSTRAINT FK_ComprobantesElectronicos_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ComprobantesElectronicos_TiposDocumentoComprobanteSuperMaestro_Codigo')
        ALTER TABLE dbo.ComprobantesElectronicos WITH CHECK ADD CONSTRAINT FK_ComprobantesElectronicos_TiposDocumentoComprobanteSuperMaestro_Codigo
            FOREIGN KEY (CodigoTipoComprobante) REFERENCES dbo.TiposDocumentoComprobanteSuperMaestro (CodigoSunat);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_ComprobantesDetalle_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.ComprobantesDetalle WITH CHECK ADD CONSTRAINT FK_ComprobantesDetalle_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_NegociosSuscripcionPago_MonedasSuperMaestro_CodigoMoneda')
        ALTER TABLE dbo.NegociosSuscripcionPago WITH CHECK ADD CONSTRAINT FK_NegociosSuscripcionPago_MonedasSuperMaestro_CodigoMoneda
            FOREIGN KEY (CodigoMoneda) REFERENCES dbo.MonedasSuperMaestro (Codigo);

    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Reservas') AND name = N'IX_Reservas_CodigoMoneda')
        CREATE NONCLUSTERED INDEX IX_Reservas_CodigoMoneda ON dbo.Reservas (CodigoMoneda);

    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Pagos') AND name = N'IX_Pagos_CodigoMoneda')
        CREATE NONCLUSTERED INDEX IX_Pagos_CodigoMoneda ON dbo.Pagos (CodigoMoneda);

    IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.ComprobantesElectronicos') AND name = N'IX_ComprobantesElectronicos_CodigosCanonicos')
        CREATE NONCLUSTERED INDEX IX_ComprobantesElectronicos_CodigosCanonicos
            ON dbo.ComprobantesElectronicos (NegocioId, CodigoTipoComprobante, CodigoMoneda);

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
    DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
    DECLARE @ErrorState INT = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Protege las monedas habilitadas por negocio cuando ya poseen movimientos.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Maestros_Monedas_Actualizar
    @NegocioId INT,
    @Id INT,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoMoneda NVARCHAR(10);

        SELECT @CodigoMoneda = m.Codigo
        FROM dbo.Monedas m
        WHERE m.Id = @Id
          AND m.NegocioId = @NegocioId;

        IF @CodigoMoneda IS NULL
            RAISERROR('No se encontro la moneda para actualizar.', 16, 1);

        IF @Activo = 0
        BEGIN
            IF EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId AND CodigoMoneda = @CodigoMoneda AND Activo = 1)
                RAISERROR('No se puede inactivar la moneda predeterminada del club.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Reservas r
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE s.NegocioId = @NegocioId
                  AND r.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen reservas registradas con ese codigo.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Pagos p
                INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE s.NegocioId = @NegocioId
                  AND p.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen pagos registrados con ese codigo.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.ComprobantesElectronicos ce
                WHERE ce.NegocioId = @NegocioId
                  AND ce.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen comprobantes emitidos con ese codigo.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Tarifas t
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = t.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE s.NegocioId = @NegocioId
                  AND t.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen tarifas configuradas con ese codigo.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.TarifaFeriado tf
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = tf.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                WHERE s.NegocioId = @NegocioId
                  AND tf.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen tarifas de feriado con ese codigo.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Cupones c
                WHERE c.NegocioId = @NegocioId
                  AND c.CodigoMoneda = @CodigoMoneda
            )
                RAISERROR('No se puede inactivar la moneda porque existen cupones registrados con ese codigo.', 16, 1);
        END;

        UPDATE dbo.Monedas
        SET Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'MAESTROS',
            @Accion = N'EDIT',
            @Entidad = N'Moneda',
            @EntidadId = @Id,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Expone y persiste el codigo canonico de moneda en la configuracion del negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Obtener
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            n.Id,
            n.NombreComercial,
            n.RazonSocial,
            COALESCE(NULLIF(n.TipoDocumentoFiscal, N''), N'1') AS TipoDocumentoFiscal,
            COALESCE(NULLIF(n.NumeroDocumentoFiscal, N''), n.DocumentoFiscal) AS NumeroDocumentoFiscal,
            n.DireccionFiscal,
            CAST(NULL AS INT) AS CompatibilidadMonedaReservada,
            n.CodigoUbigeo,
            CAST(COALESCE(n.PoliticaConfirmacionPago, 0) AS TINYINT) AS PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo,
            CAST(COALESCE(n.EmisionComprobantesElectronicos, 0) AS BIT) AS EmisionComprobantesElectronicos,
            CAST(COALESCE(n.EnviarComprobanteAutomatico, 0) AS BIT) AS EnviarComprobanteAutomatico,
            CAST(COALESCE(n.EmisionReciboInterno, 0) AS BIT) AS EmisionReciboInterno,
            CAST(COALESCE(n.PorcentajeIgv, 18) AS INT) AS PorcentajeIgv,
            n.LogoUrl,
            CAST(COALESCE(n.PermitirModificarPrecioReserva, 0) AS BIT) AS PermitirModificarPrecioReserva,
            CAST(COALESCE(n.CancelacionAutomaticaNoConfirmada, 0) AS BIT) AS CancelacionAutomaticaNoConfirmada,
            CAST(COALESCE(n.MinutosCancelacionNoConfirmada, 30) AS INT) AS MinutosCancelacionNoConfirmada,
            CAST(COALESCE(n.SedesPermitidas, 2) AS INT) AS SedesPermitidas,
            CAST(COALESCE(n.EspaciosPermitidos, 6) AS INT) AS EspaciosPermitidos,
            CAST(COALESCE(n.HorasMaximasReservaCliente, 1) AS INT) AS HorasMaximasReservaCliente,
            n.CodigoMoneda,
            COALESCE(ms.Simbolo, n.CodigoMoneda) AS MonedaSimbolo
        FROM dbo.Negocios n
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Codigo = n.CodigoMoneda
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Guarda la configuracion del club con moneda canonica validada por negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @CodigoMoneda NVARCHAR(10),
    @PoliticaConfirmacionPago TINYINT = 0,
    @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL,
    @EmisionComprobantesElectronicos BIT = 0,
    @EnviarComprobanteAutomatico BIT = 0,
    @EmisionReciboInterno BIT = 0,
    @PorcentajeIgv INT = 18,
    @LogoUrl NVARCHAR(500) = NULL,
    @PermitirModificarPrecioReserva BIT = 0,
    @CancelacionAutomaticaNoConfirmada BIT = 0,
    @MinutosCancelacionNoConfirmada INT = NULL,
    @HorasMaximasReservaCliente INT = 1,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @DireccionFiscalNormalizada NVARCHAR(250);
        DECLARE @CodigoUbigeoNormalizado CHAR(6);
        DECLARE @PorcentajeAdelantoNormalizado DECIMAL(5,2);
        DECLARE @LogoUrlNormalizado NVARCHAR(500);

        SET @TipoDocumentoFiscal = NULLIF(UPPER(LTRIM(RTRIM(@TipoDocumentoFiscal))), N'');
        SET @DireccionFiscalNormalizada = NULLIF(LTRIM(RTRIM(@DireccionFiscal)), N'');
        SET @CodigoUbigeoNormalizado = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @PorcentajeAdelantoNormalizado = @PorcentajeAdelantoMinimo;
        SET @LogoUrlNormalizado = NULLIF(LTRIM(RTRIM(@LogoUrl)), N'');
        SET @CodigoMoneda = NULLIF(UPPER(LTRIM(RTRIM(@CodigoMoneda))), N'');

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.Monedas m
            INNER JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = m.Codigo AND msm.Activo = 1
            WHERE m.NegocioId = @NegocioId
              AND m.Codigo = @CodigoMoneda
              AND m.Activo = 1
        )
            RAISERROR('La moneda seleccionada no esta habilitada para este club.', 16, 1);

        IF @TipoDocumentoFiscal IS NULL
            RAISERROR('El tipo de documento SUNAT es obligatorio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = @TipoDocumentoFiscal AND t.Activo = 1)
            RAISERROR('El tipo de documento SUNAT no es valido.', 16, 1);

        IF @DireccionFiscalNormalizada IS NULL
            SET @CodigoUbigeoNormalizado = NULL;

        IF @DireccionFiscalNormalizada IS NOT NULL AND @CodigoUbigeoNormalizado IS NULL
            RAISERROR('Cuando se registra direccion fiscal, el distrito es obligatorio.', 16, 1);

        IF @CodigoUbigeoNormalizado IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeoNormalizado AND Activo = 1)
            RAISERROR('El codigo de ubigeo no existe.', 16, 1);

        IF @PoliticaConfirmacionPago NOT IN (0, 1, 2)
            RAISERROR('La politica de confirmacion no es valida.', 16, 1);

        IF @PoliticaConfirmacionPago = 1
        BEGIN
            IF @PorcentajeAdelantoNormalizado IS NULL OR @PorcentajeAdelantoNormalizado < 1 OR @PorcentajeAdelantoNormalizado > 100
                RAISERROR('Para exigir adelanto, el porcentaje minimo debe ser entero entre 1 y 100.', 16, 1);

            IF @PorcentajeAdelantoNormalizado <> FLOOR(@PorcentajeAdelantoNormalizado)
                RAISERROR('El porcentaje minimo de adelanto no admite decimales.', 16, 1);
        END
        ELSE
            SET @PorcentajeAdelantoNormalizado = NULL;

        IF @PorcentajeIgv IS NULL OR @PorcentajeIgv < 0 OR @PorcentajeIgv > 100
            RAISERROR('El porcentaje de IGV debe estar entre 0 y 100.', 16, 1);

        IF @CancelacionAutomaticaNoConfirmada = 1
        BEGIN
            IF @MinutosCancelacionNoConfirmada IS NULL OR @MinutosCancelacionNoConfirmada < 5 OR @MinutosCancelacionNoConfirmada > 1440
                RAISERROR('El tiempo de cancelacion automatica debe estar entre 5 y 1440 minutos.', 16, 1);
        END
        ELSE
            SET @MinutosCancelacionNoConfirmada = NULL;

        IF @HorasMaximasReservaCliente IS NULL OR @HorasMaximasReservaCliente < 1 OR @HorasMaximasReservaCliente > 12
            RAISERROR('La hora(s) maxima de reserva por cliente debe estar entre 1 y 12.', 16, 1);

        UPDATE n
        SET n.NombreComercial = @NombreComercial,
            n.RazonSocial = NULLIF(@RazonSocial, N''),
            n.TipoDocumentoFiscal = @TipoDocumentoFiscal,
            n.NumeroDocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.DireccionFiscal = @DireccionFiscalNormalizada,
            n.CodigoUbigeo = @CodigoUbigeoNormalizado,
            n.DocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.CodigoMoneda = @CodigoMoneda,
            n.PoliticaConfirmacionPago = @PoliticaConfirmacionPago,
            n.PorcentajeAdelantoMinimo = @PorcentajeAdelantoNormalizado,
            n.EmisionComprobantesElectronicos = @EmisionComprobantesElectronicos,
            n.EnviarComprobanteAutomatico = @EnviarComprobanteAutomatico,
            n.EmisionReciboInterno = @EmisionReciboInterno,
            n.PorcentajeIgv = @PorcentajeIgv,
            n.LogoUrl = @LogoUrlNormalizado,
            n.PermitirModificarPrecioReserva = @PermitirModificarPrecioReserva,
            n.CancelacionAutomaticaNoConfirmada = @CancelacionAutomaticaNoConfirmada,
            n.MinutosCancelacionNoConfirmada = @MinutosCancelacionNoConfirmada,
            n.HorasMaximasReservaCliente = @HorasMaximasReservaCliente
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80) = CONVERT(NVARCHAR(80), @NegocioId);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Registra pagos conservando la moneda canonica de la reserva.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Crear
    @NegocioId INT,
    @ReservaId INT,
    @FechaPago DATETIME2,
    @Monto DECIMAL(10,2),
    @FormaPago INT,
    @NumeroOperacion NVARCHAR(50) = NULL,
    @Observacion NVARCHAR(300) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @Monto <= 0
            RAISERROR('El monto debe ser mayor que cero.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FormasPago WHERE Id = @FormaPago AND NegocioId = @NegocioId AND Activo = 1)
            RAISERROR('La forma de pago no es valida para el negocio.', 16, 1);

        DECLARE @TotalReserva DECIMAL(10,2);
        DECLARE @PagadoActual DECIMAL(10,2);
        DECLARE @NuevoPagado DECIMAL(10,2);
        DECLARE @SaldoPendiente DECIMAL(10,2);
        DECLARE @PoliticaConfirmacionPago TINYINT = 0;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL;
        DECLARE @MontoMinimoAdelanto DECIMAL(10,2) = NULL;
        DECLARE @CodigoMoneda NVARCHAR(10);

        SELECT
            @TotalReserva = r.Total,
            @CodigoMoneda = r.CodigoMoneda,
            @PoliticaConfirmacionPago = ISNULL(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE r.Id = @ReservaId
          AND s.NegocioId = @NegocioId;

        IF @TotalReserva IS NULL OR @CodigoMoneda IS NULL
            RAISERROR('Reserva invalida o sin moneda canonica para el negocio.', 16, 1);

        SELECT @PagadoActual = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        WHERE p.ReservaId = @ReservaId;

        SET @SaldoPendiente = @TotalReserva - @PagadoActual;

        IF @SaldoPendiente <= 0
            RAISERROR('La reserva ya esta pagada al 100%. No se pueden registrar mas pagos.', 16, 1);

        IF @Monto > @SaldoPendiente
            RAISERROR('El pago excede el saldo pendiente de la reserva.', 16, 1);

        SET @NuevoPagado = @PagadoActual + @Monto;
        IF @NuevoPagado > @TotalReserva
            RAISERROR('El pago excede el total de la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 2 AND @NuevoPagado < @TotalReserva
            RAISERROR('La configuracion del negocio exige pago total (100%) para confirmar la reserva.', 16, 1);

        IF @PoliticaConfirmacionPago = 1
        BEGIN
            SET @PorcentajeAdelantoMinimo = ISNULL(@PorcentajeAdelantoMinimo, 0);
            IF @PorcentajeAdelantoMinimo > 0
            BEGIN
                SET @MontoMinimoAdelanto = ROUND((@TotalReserva * @PorcentajeAdelantoMinimo) / 100.0, 2);
                IF @NuevoPagado < @MontoMinimoAdelanto AND @NuevoPagado < @TotalReserva
                    RAISERROR('El pago acumulado no alcanza el adelanto minimo configurado para confirmar la reserva.', 16, 1);
            END
        END;

        BEGIN TRANSACTION;

        DECLARE @NumeroPorNegocio INT;
        EXEC dbo.Sp_NegocioCorrelativos_ObtenerSiguiente @NegocioId, N'PAGO', @Usuario, @NumeroPorNegocio OUTPUT;

        INSERT INTO dbo.Pagos
        (
            NumeroPorNegocio, ReservaId, FechaPago, Monto, CodigoMoneda, FormaPago, NumeroOperacion, Observacion,
            FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NumeroPorNegocio, @ReservaId, @FechaPago, @Monto, @CodigoMoneda, @FormaPago, @NumeroOperacion, @Observacion,
            SYSUTCDATETIME(), @Usuario
        );

        UPDATE r
        SET Adelanto = @NuevoPagado,
            Saldo = r.Total - @NuevoPagado,
            Estado = CASE
                        WHEN r.Total - @NuevoPagado <= 0 THEN 4
                        WHEN @NuevoPagado > 0 THEN 2
                        ELSE r.Estado
                     END,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        FROM dbo.Reservas r
        WHERE r.Id = @ReservaId;

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAuditoria NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'PAGOS',
            @Accion = N'CREATE',
            @Entidad = N'Pago',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
