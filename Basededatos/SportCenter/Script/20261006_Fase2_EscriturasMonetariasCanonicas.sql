
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Fase 2. Persiste codigos monetarios canonicos en rutas de escritura.
-- Firma:         FRANCO LARA - 06/10/2026 | Migra escrituras canonicas, corrige parametros de auditoria, valida SP publicados y cierra columnas transaccionales como obligatorias.
-- Requiere:      20261006_CatalogosCanonicos_MonedasYComprobantes.sql
-- Despliegue:    Publicar primero los SP de escritura enumerados al final de este encabezado
--                y ejecutar despues este script para validar/backfillear y cerrar NOT NULL.
-- SP requeridos: Sp_Reservas_Crear, Sp_Reservas_Actualizar, Sp_Pagos_Crear,
--                Sp_Espacios_Crear, Sp_Espacios_Actualizar, Sp_Cupones_Crear,
--                Sp_SolicitudesPublicas_ConvertirAReserva, Sp_Comprobantes_Crear,
--                Sp_NegociosSuscripcionPago_Registrar, Sp_Maestros_Monedas_Crear
--                y Sp_Maestros_Monedas_Actualizar.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Cupones_Crear
    @NegocioId INT,
    @SedeId INT = NULL,
    @EspacioDeportivoId INT = NULL,
    @CodigoCupon NVARCHAR(30),
    @Nombre NVARCHAR(150),
    @TipoDescuento NVARCHAR(20),
    @ValorDescuento DECIMAL(10,2),
    @CantidadMaxUsos INT,
    @FechaInicio DATE,
    @FechaFin DATE,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoMoneda NVARCHAR(10);

        SELECT @CodigoMoneda = n.CodigoMoneda
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @CodigoMoneda IS NULL
            RAISERROR('El negocio debe configurar una moneda valida antes de crear cupones.', 16, 1);

        SET @CodigoCupon = UPPER(LTRIM(RTRIM(@CodigoCupon)));
        IF @CodigoCupon = N'' RAISERROR('El codigo de cupon es obligatorio.', 16, 1);
        IF @FechaFin < @FechaInicio RAISERROR('La fecha fin no puede ser menor a la fecha inicio.', 16, 1);
        IF @CantidadMaxUsos <= 0 RAISERROR('La cantidad maxima de usos debe ser mayor a cero.', 16, 1);
        IF @TipoDescuento NOT IN (N'PORCENTAJE', N'MONTO_FIJO') RAISERROR('Tipo de descuento no valido.', 16, 1);
        IF @TipoDescuento = N'PORCENTAJE' AND (@ValorDescuento <= 0 OR @ValorDescuento > 100) RAISERROR('El porcentaje debe estar entre 0.01 y 100.', 16, 1);
        IF @TipoDescuento = N'MONTO_FIJO' AND @ValorDescuento <= 0 RAISERROR('El monto fijo debe ser mayor a cero.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.Cupones WHERE NegocioId = @NegocioId AND CodigoCupon = @CodigoCupon)
            RAISERROR('El codigo de cupon ya existe para este negocio.', 16, 1);

        DECLARE @NumeroPorNegocio INT;
        EXEC dbo.Sp_NegocioCorrelativos_ObtenerSiguiente @NegocioId, N'CUPON', @Usuario, @NumeroPorNegocio OUTPUT;

        INSERT INTO dbo.Cupones
        (
            NumeroPorNegocio, NegocioId, SedeId, EspacioDeportivoId, CodigoCupon, Nombre, TipoDescuento, ValorDescuento,
            CodigoMoneda, CantidadMaxUsos, CantidadUsosActuales, FechaInicio, FechaFin, Activo, FechaRegistro, UsuarioCreacion
        )
        VALUES
        (
            @NumeroPorNegocio, @NegocioId, @SedeId, @EspacioDeportivoId, @CodigoCupon, @Nombre, @TipoDescuento, @ValorDescuento,
            @CodigoMoneda, @CantidadMaxUsos, 0, @FechaInicio, @FechaFin, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT = SCOPE_IDENTITY();
        DECLARE @EntidadIdAuditoria NVARCHAR(80) = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CUPONES',
            @Accion = N'CREATE',
            @Entidad = N'Cupon',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        SELECT @Id;
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
-- Description:   Completa datos canonicos y activa su obligatoriedad transaccional.
-- =============================================
SET XACT_ABORT ON;
BEGIN TRANSACTION;

BEGIN TRY
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Reservas_Crear')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Reservas_Crear con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Reservas_Actualizar')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Reservas_Actualizar con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Pagos_Crear')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Pagos_Crear con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Espacios_Crear')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Espacios_Crear con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Espacios_Actualizar')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Espacios_Actualizar con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Cupones_Crear')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Cupones_Crear con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_SolicitudesPublicas_ConvertirAReserva')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_SolicitudesPublicas_ConvertirAReserva con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Comprobantes_Crear')), N'') NOT LIKE N'%CodigoTipoComprobante%'
       OR COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Comprobantes_Crear')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Comprobantes_Crear con codigos canonicos.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_NegociosSuscripcionPago_Registrar')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_NegociosSuscripcionPago_Registrar con moneda canonica.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Maestros_Monedas_Actualizar')), N'') NOT LIKE N'%TarifaFeriado%'
       OR COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Maestros_Monedas_Actualizar')), N'') NOT LIKE N'%CodigoMoneda%'
        RAISERROR('Fase 2 incompleta: publique Sp_Maestros_Monedas_Actualizar con proteccion de movimientos.', 16, 1);
    IF COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Maestros_Monedas_Crear')), N'') = N''
       OR COALESCE(OBJECT_DEFINITION(OBJECT_ID(N'dbo.Sp_Maestros_Monedas_Crear')), N'') LIKE N'%Solo se permite una moneda por negocio%'
        RAISERROR('Fase 2 incompleta: publique Sp_Maestros_Monedas_Crear con soporte multimoneda.', 16, 1);

    UPDATE r
       SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Reservas r
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
    WHERE r.CodigoMoneda IS NULL;

    UPDATE p SET CodigoMoneda = r.CodigoMoneda
    FROM dbo.Pagos p INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
    WHERE p.CodigoMoneda IS NULL;

    UPDATE t SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Tarifas t
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = t.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
    WHERE t.CodigoMoneda IS NULL;

    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
        UPDATE tf SET CodigoMoneda = n.CodigoMoneda
        FROM dbo.TarifaFeriado tf
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = tf.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Negocios n ON n.Id = s.NegocioId
        WHERE tf.CodigoMoneda IS NULL;

    UPDATE c SET CodigoMoneda = n.CodigoMoneda
    FROM dbo.Cupones c INNER JOIN dbo.Negocios n ON n.Id = c.NegocioId
    WHERE c.CodigoMoneda IS NULL;

    UPDATE cu SET CodigoMoneda = r.CodigoMoneda
    FROM dbo.CuponesUso cu INNER JOIN dbo.Reservas r ON r.Id = cu.ReservaId
    WHERE cu.CodigoMoneda IS NULL;

    UPDATE ce
       SET CodigoMoneda = r.CodigoMoneda,
           CodigoTipoComprobante = ntd.CodigoSunat
    FROM dbo.ComprobantesElectronicos ce
    INNER JOIN dbo.Reservas r ON r.Id = ce.ReservaId
    INNER JOIN dbo.NegociosTiposDocumentoComprobante ntd
        ON ntd.Id = ce.TipoComprobante AND ntd.NegocioId = ce.NegocioId
    WHERE ce.CodigoMoneda IS NULL OR ce.CodigoTipoComprobante IS NULL;

    UPDATE cd SET CodigoMoneda = ce.CodigoMoneda
    FROM dbo.ComprobantesDetalle cd
    INNER JOIN dbo.ComprobantesElectronicos ce ON ce.Id = cd.ComprobanteElectronicoId
    WHERE cd.CodigoMoneda IS NULL;

    UPDATE nsp SET CodigoMoneda = UPPER(LTRIM(RTRIM(nsp.Moneda)))
    FROM dbo.NegociosSuscripcionPago nsp
    WHERE nsp.CodigoMoneda IS NULL;

    IF EXISTS (SELECT 1 FROM dbo.Reservas WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen reservas sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.Pagos WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen pagos sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.Tarifas WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen tarifas sin CodigoMoneda.', 16, 1);
    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL AND EXISTS (SELECT 1 FROM dbo.TarifaFeriado WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen tarifas de feriado sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.Cupones WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen cupones sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.CuponesUso WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen usos de cupon sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.ComprobantesElectronicos WHERE CodigoMoneda IS NULL OR CodigoTipoComprobante IS NULL)
        RAISERROR('Fase 2 incompleta: existen comprobantes sin codigos canonicos.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.ComprobantesDetalle WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen detalles de comprobante sin CodigoMoneda.', 16, 1);
    IF EXISTS (SELECT 1 FROM dbo.NegociosSuscripcionPago WHERE CodigoMoneda IS NULL)
        RAISERROR('Fase 2 incompleta: existen cobros de suscripcion sin CodigoMoneda.', 16, 1);

    ALTER TABLE dbo.Reservas ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.Pagos ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.Tarifas ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    IF OBJECT_ID(N'dbo.TarifaFeriado', N'U') IS NOT NULL
        ALTER TABLE dbo.TarifaFeriado ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.Cupones ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.CuponesUso ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.ComprobantesElectronicos ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.ComprobantesElectronicos ALTER COLUMN CodigoTipoComprobante NVARCHAR(4) NOT NULL;
    ALTER TABLE dbo.ComprobantesDetalle ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;
    ALTER TABLE dbo.NegociosSuscripcionPago ALTER COLUMN CodigoMoneda NVARCHAR(10) NOT NULL;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
    DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
    DECLARE @ErrorState INT = ERROR_STATE();
    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH
GO
