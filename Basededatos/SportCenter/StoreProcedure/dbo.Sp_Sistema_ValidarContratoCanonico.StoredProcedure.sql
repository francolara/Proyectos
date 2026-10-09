
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Valida continuamente contratos canonicos de moneda, comprobantes, deporte y suelo.
-- Firma:         FRANCO LARA - 08/10/2026 | Acepta negocios pendientes de configuracion sin CodigoMoneda.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Sistema_ValidarContratoCanonico
    @ValidarDatos BIT = 1
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF COL_LENGTH(N'dbo.Negocios', N'CodigoMoneda') IS NULL
           OR COL_LENGTH(N'dbo.Reservas', N'CodigoMoneda') IS NULL
           OR COL_LENGTH(N'dbo.Pagos', N'CodigoMoneda') IS NULL
           OR COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoMoneda') IS NULL
           OR COL_LENGTH(N'dbo.ComprobantesElectronicos', N'CodigoTipoComprobante') IS NULL
           OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteSuperId') IS NULL
           OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloSuperId') IS NULL
            RAISERROR('Contrato canonico incompleto: faltan columnas obligatorias.', 16, 1);

        IF COL_LENGTH(N'dbo.Negocios', N'MonedaId') IS NOT NULL
           OR COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoMoneda') IS NOT NULL
           OR COL_LENGTH(N'dbo.ComprobantesElectronicos', N'TipoComprobante') IS NOT NULL
           OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoDeporteId') IS NOT NULL
           OR COL_LENGTH(N'dbo.EspaciosDeportivos', N'TipoSueloId') IS NOT NULL
            RAISERROR('Contrato canonico incompleto: aun existen columnas heredadas.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM sys.columns c
            INNER JOIN sys.tables t ON t.object_id = c.object_id
            INNER JOIN sys.schemas s ON s.schema_id = t.schema_id
            WHERE s.name = N'dbo'
              AND c.name IN (N'CodigoMoneda', N'CodigoTipoComprobante')
              AND t.name IN
              (
                  N'Reservas', N'Pagos', N'Tarifas', N'TarifaFeriado',
                  N'Cupones', N'CuponesUso', N'ComprobantesElectronicos',
                  N'ComprobantesDetalle', N'NegociosSuscripcionPago'
              )
              AND c.is_nullable = 1
        )
            RAISERROR('Contrato canonico incompleto: existen columnas canonicas que permiten NULL.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM sys.columns c
            WHERE c.object_id = OBJECT_ID(N'dbo.EspaciosDeportivos')
              AND c.name IN (N'TipoDeporteSuperId', N'TipoSueloSuperId')
              AND c.is_nullable = 1
        )
            RAISERROR('Contrato canonico incompleto: deporte o suelo global permiten NULL.', 16, 1);

        IF @ValidarDatos = 1
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.EspaciosDeportivos e
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                LEFT JOIN dbo.TiposDeporteSuperMaestro tds ON tds.Id = e.TipoDeporteSuperId
                LEFT JOIN dbo.TiposSueloSuperMaestro tss ON tss.Id = e.TipoSueloSuperId
                LEFT JOIN dbo.TiposDeporte td
                    ON td.NegocioId = s.NegocioId
                   AND td.TipoDeporteSuperId = e.TipoDeporteSuperId
                LEFT JOIN dbo.TiposSuelo ts
                    ON ts.NegocioId = s.NegocioId
                   AND ts.TipoSueloSuperId = e.TipoSueloSuperId
                WHERE tds.Id IS NULL
                   OR tss.Id IS NULL
                   OR td.Id IS NULL
                   OR td.Activo <> 1
                   OR ts.Id IS NULL
                   OR ts.Activo <> 1
            )
                RAISERROR('Integridad canonica invalida: hay espacios con deporte o suelo global no asociado a su negocio.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Negocios n
                LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = n.CodigoMoneda
                LEFT JOIN dbo.Monedas m
                    ON m.NegocioId = n.Id
                   AND m.Codigo = n.CodigoMoneda
                   AND m.Activo = 1
                WHERE NULLIF(LTRIM(RTRIM(n.CodigoMoneda)), N'') IS NOT NULL
                  AND (msm.Codigo IS NULL OR m.Id IS NULL)
            )
                RAISERROR('Integridad canonica invalida: hay negocios con una moneda configurada que no es global o no esta activa localmente.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Reservas r
                INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
                INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
                LEFT JOIN dbo.MonedasSuperMaestro msm ON msm.Codigo = r.CodigoMoneda
                LEFT JOIN dbo.Monedas m
                    ON m.NegocioId = s.NegocioId
                   AND m.Codigo = r.CodigoMoneda
                WHERE NULLIF(LTRIM(RTRIM(r.CodigoMoneda)), N'') IS NULL
                   OR msm.Codigo IS NULL
                   OR m.Id IS NULL
            )
                RAISERROR('Integridad canonica invalida: hay reservas con moneda inexistente para su negocio.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.Pagos p
                INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
                WHERE p.CodigoMoneda <> r.CodigoMoneda
            )
                RAISERROR('Integridad historica invalida: la moneda de un pago difiere de su reserva.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CuponesUso cu
                INNER JOIN dbo.Reservas r ON r.Id = cu.ReservaId
                WHERE cu.CodigoMoneda <> r.CodigoMoneda
            )
                RAISERROR('Integridad historica invalida: la moneda de un uso de cupon difiere de su reserva.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.ComprobantesElectronicos ce
                INNER JOIN dbo.Reservas r ON r.Id = ce.ReservaId
                LEFT JOIN dbo.TiposDocumentoComprobanteSuperMaestro td
                    ON td.CodigoSunat = ce.CodigoTipoComprobante
                WHERE ce.CodigoMoneda <> r.CodigoMoneda
                   OR NULLIF(LTRIM(RTRIM(ce.CodigoTipoComprobante)), N'') IS NULL
                   OR td.CodigoSunat IS NULL
            )
                RAISERROR('Integridad canonica invalida: hay comprobantes con moneda o codigo SUNAT inconsistente.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.ComprobantesDetalle cd
                INNER JOIN dbo.ComprobantesElectronicos ce ON ce.Id = cd.ComprobanteElectronicoId
                WHERE cd.CodigoMoneda <> ce.CodigoMoneda
            )
                RAISERROR('Integridad historica invalida: la moneda del detalle difiere de su comprobante.', 16, 1);

            IF EXISTS
            (
                SELECT
                    ce.NegocioId,
                    ce.CodigoTipoComprobante,
                    ce.Serie,
                    ce.Numero
                FROM dbo.ComprobantesElectronicos ce
                GROUP BY
                    ce.NegocioId,
                    ce.CodigoTipoComprobante,
                    ce.Serie,
                    ce.Numero
                HAVING COUNT_BIG(*) > 1
            )
                RAISERROR('Integridad canonica invalida: existen correlativos de comprobante duplicados.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.indexes i
            WHERE i.object_id = OBJECT_ID(N'dbo.ComprobantesElectronicos')
              AND i.name = N'UX_ComprobantesElectronicos_Negocio_CodigoTipo_Serie_Numero'
              AND i.is_unique = 1
              AND i.is_disabled = 0
        )
            RAISERROR('Contrato canonico incompleto: falta el indice unico de comprobantes.', 16, 1);

        DECLARE @ClavesCanonicas TABLE
        (
            Nombre SYSNAME NOT NULL PRIMARY KEY
        );

        INSERT INTO @ClavesCanonicas (Nombre)
        VALUES
            (N'FK_Negocios_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_Reservas_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_Pagos_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_Tarifas_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_Cupones_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_CuponesUso_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_ComprobantesElectronicos_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_ComprobantesElectronicos_TiposDocumentoComprobanteSuperMaestro_Codigo'),
            (N'FK_ComprobantesDetalle_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_NegociosSuscripcionPago_MonedasSuperMaestro_CodigoMoneda'),
            (N'FK_EspaciosDeportivos_TiposDeporteSuperMaestro_TipoDeporteSuperId'),
            (N'FK_EspaciosDeportivos_TiposSueloSuperMaestro_TipoSueloSuperId');

        IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
            INSERT INTO @ClavesCanonicas (Nombre)
            VALUES (N'FK_TarifaFeriado_MonedasSuperMaestro_CodigoMoneda');

        IF EXISTS
        (
            SELECT 1
            FROM @ClavesCanonicas cc
            LEFT JOIN sys.foreign_keys fk ON fk.name = cc.Nombre
            WHERE fk.object_id IS NULL
        )
            RAISERROR('Contrato canonico incompleto: faltan claves foraneas canonicas.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM @ClavesCanonicas cc
            INNER JOIN sys.foreign_keys fk ON fk.name = cc.Nombre
            WHERE fk.is_disabled = 1
               OR fk.is_not_trusted = 1
        )
            RAISERROR('Contrato canonico incompleto: existen claves foraneas deshabilitadas o no confiables.', 16, 1);

        IF OBJECT_ID(N'dbo.Sp_ConfiguracionClub_Actualizar', N'P') IS NULL
           OR OBJECT_ID(N'dbo.Sp_Comprobantes_Crear', N'P') IS NULL
           OR OBJECT_ID(N'dbo.Sp_Espacios_Crear', N'P') IS NULL
           OR OBJECT_ID(N'dbo.Sp_Espacios_Actualizar', N'P') IS NULL
            RAISERROR('Contrato canonico incompleto: faltan procedimientos principales.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM sys.parameters p
            WHERE p.object_id IN
            (
                OBJECT_ID(N'dbo.Sp_ConfiguracionClub_Actualizar'),
                OBJECT_ID(N'dbo.Sp_Comprobantes_Crear'),
                OBJECT_ID(N'dbo.Sp_Espacios_Crear'),
                OBJECT_ID(N'dbo.Sp_Espacios_Actualizar')
            )
              AND p.name IN (N'@MonedaId', N'@TipoMoneda', N'@TipoComprobante', N'@TipoDeporteId', N'@TipoSueloId')
        )
            RAISERROR('Contrato canonico incompleto: un procedimiento aun expone parametros heredados.', 16, 1);

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.parameters p
            WHERE p.object_id = OBJECT_ID(N'dbo.Sp_ConfiguracionClub_Actualizar')
              AND p.name = N'@CodigoMoneda'
        )
           OR NOT EXISTS
        (
            SELECT 1
            FROM sys.parameters p
            WHERE p.object_id = OBJECT_ID(N'dbo.Sp_Comprobantes_Crear')
              AND p.name = N'@CodigoDocumentoComprobante'
        )
           OR NOT EXISTS
        (
            SELECT 1
            FROM sys.parameters p
            WHERE p.object_id = OBJECT_ID(N'dbo.Sp_Espacios_Crear')
              AND p.name = N'@TipoDeporteSuperId'
        )
           OR NOT EXISTS
        (
            SELECT 1
            FROM sys.parameters p
            WHERE p.object_id = OBJECT_ID(N'dbo.Sp_Espacios_Crear')
              AND p.name = N'@TipoSueloSuperId'
        )
            RAISERROR('Contrato canonico incompleto: faltan parametros canonicos principales.', 16, 1);

        DECLARE @Reportes TABLE
        (
            Nombre SYSNAME NOT NULL PRIMARY KEY
        );

        INSERT INTO @Reportes (Nombre)
        VALUES
            (N'Sp_Reportes_IngresosPorDia'),
            (N'Sp_Reportes_ReservasPorDia'),
            (N'Sp_Reportes_OcupacionPorEspacio'),
            (N'Sp_Reportes_ResumenOperativo'),
            (N'Sp_Reportes_ResumenCobranza'),
            (N'Sp_Reportes_DetallePagos'),
            (N'Sp_Reportes_DetalleReservas');

        IF EXISTS
        (
            SELECT 1
            FROM @Reportes r
            LEFT JOIN sys.procedures p
                ON p.schema_id = SCHEMA_ID(N'dbo')
               AND p.name = r.Nombre
            LEFT JOIN sys.parameters prm
                ON prm.object_id = p.object_id
               AND prm.name = N'@CodigoMoneda'
            WHERE p.object_id IS NULL
               OR prm.parameter_id IS NULL
        )
            RAISERROR('Contrato canonico incompleto: un reporte no exige CodigoMoneda.', 16, 1);

        SELECT
            CAST(1 AS BIT) AS EsValido,
            CAST(N'CANONICO_MAESTROS_V2' AS NVARCHAR(40)) AS VersionContrato,
            SYSUTCDATETIME() AS FechaValidacionUtc;
    END TRY
    BEGIN CATCH
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
GO
