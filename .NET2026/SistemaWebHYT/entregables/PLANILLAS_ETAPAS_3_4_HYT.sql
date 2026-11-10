-- Firma: FRANCO LARA | Create date: 09/11/2026 | Pagos periódicos, cierre, ceses, obligaciones y reportes de Planillas.
-- Paquete incremental. Ejecutar con sqlcmd -b -f 65001 después de PLANILLAS_HYT.sql.
USE [SistemaWebHYT];
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO

-- Fuente: Script/034_PeriodicidadPagoProveedor.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Agrega frecuencia y periodo de cobertura a los abonos de proveedor.
-- =============================================
USE [SistemaWebHYT];
GO
IF COL_LENGTH(N'dbo.PagoProveedorHYT', N'Frecuencia') IS NULL
BEGIN
    ALTER TABLE dbo.PagoProveedorHYT ADD Frecuencia VARCHAR(10) NULL;
    ALTER TABLE dbo.PagoProveedorHYT ADD PeriodoDesde DATE NULL;
    ALTER TABLE dbo.PagoProveedorHYT ADD PeriodoHasta DATE NULL;
END;
GO
IF OBJECT_ID(N'dbo.CK_PagoProveedorHYT_Periodo', N'C') IS NULL
BEGIN
    ALTER TABLE dbo.PagoProveedorHYT WITH CHECK ADD CONSTRAINT CK_PagoProveedorHYT_Periodo CHECK (
        (Frecuencia IS NULL AND PeriodoDesde IS NULL AND PeriodoHasta IS NULL)
        OR (Frecuencia IS NOT NULL AND Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
            AND PeriodoDesde IS NOT NULL AND PeriodoHasta IS NOT NULL AND PeriodoDesde <= PeriodoHasta));
END;
GO

-- Fuente: StoreProcedure/spu_PagoProveedor_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra abonos idempotentes con frecuencia y periodo de cobertura.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PagoProveedor_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @CostoProveedorId BIGINT,
    @Importe DECIMAL(20,4),
    @FechaPago DATE,
    @Referencia NVARCHAR(100),
    @Frecuencia VARCHAR(10) = NULL,
    @PeriodoDesde DATE = NULL,
    @PeriodoHasta DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar pagos a proveedores.', 1;

        SET @Referencia = TRIM(@Referencia);
        IF (@Frecuencia IS NULL AND (@PeriodoDesde IS NOT NULL OR @PeriodoHasta IS NOT NULL))
           OR (@Frecuencia IS NOT NULL AND (@Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
               OR @PeriodoDesde IS NULL OR @PeriodoHasta IS NULL OR @PeriodoHasta < @PeriodoDesde))
            THROW 51000, N'Revise la frecuencia y las fechas del periodo.', 1;
        IF @OperacionId IS NULL OR @FechaPago IS NULL OR NULLIF(@Referencia, N'') IS NULL
           OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
            THROW 51000, N'Indique fecha, referencia e importe positivo con hasta dos decimales.', 1;

        BEGIN TRANSACTION;

        -- Serializa los abonos de un mismo costo hasta confirmar el pago y su auditoria.
        DECLARE @Costo DECIMAL(18,2);
        SELECT @Costo = costo.Importe
        FROM dbo.CostoProveedorPresupuesto AS costo WITH (UPDLOCK, HOLDLOCK)
        WHERE costo.CostoProveedorId = @CostoProveedorId;

        IF @Costo IS NULL
            THROW 51000, N'El costo del proveedor no existe.', 1;

        DECLARE @PagoId BIGINT;
        SELECT @PagoId = pago.PagoProveedorId
        FROM dbo.PagoProveedorHYT AS pago
        WHERE pago.OperacionId = @OperacionId;

        IF @PagoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.PagoProveedorHYT AS pago
                WHERE pago.PagoProveedorId = @PagoId AND pago.UsuarioRegistro = @UsuarioId
                    AND pago.CostoProveedorId = @CostoProveedorId AND pago.Importe = @Importe
                    AND pago.FechaPago = @FechaPago AND pago.Referencia = @Referencia
                    AND pago.Frecuencia IS NOT DISTINCT FROM @Frecuencia
                    AND pago.PeriodoDesde IS NOT DISTINCT FROM @PeriodoDesde
                    AND pago.PeriodoHasta IS NOT DISTINCT FROM @PeriodoHasta)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Pagado DECIMAL(18,2);
            SELECT @Pagado = COALESCE(SUM(pago.Importe), 0)
            FROM dbo.PagoProveedorHYT AS pago
            WHERE pago.CostoProveedorId = @CostoProveedorId;

            IF @Importe > @Costo - @Pagado
                THROW 51000, N'El importe excede el saldo pendiente.', 1;

            INSERT INTO dbo.PagoProveedorHYT
                (CostoProveedorId, Importe, FechaPago, Referencia, OperacionId, UsuarioRegistro,
                 Frecuencia, PeriodoDesde, PeriodoHasta)
            VALUES (@CostoProveedorId, @Importe, @FechaPago, @Referencia, @OperacionId, @UsuarioId,
                    @Frecuencia, @PeriodoDesde, @PeriodoHasta);

            SET @PagoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PagoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'PagoProveedorHYT', @RegistroId, N'REGISTRAR';
        END;

        COMMIT TRANSACTION;
        SELECT @PagoId AS PagoProveedorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PagoProveedor_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta el historial de abonos con frecuencia y periodo de cobertura.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PagoProveedor_Listar
    @UsuarioId NVARCHAR(128),
    @CostoProveedorId BIGINT,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar pagos a proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        SELECT pago.PagoProveedorId, pago.Importe, pago.FechaPago, pago.Referencia,
               pago.Frecuencia, pago.PeriodoDesde, pago.PeriodoHasta,
               pago.FechaRegistro, pago.UsuarioRegistro, costo.MonedaCodigo,
               COUNT_BIG(1) OVER () AS Total
        FROM dbo.PagoProveedorHYT AS pago
        INNER JOIN dbo.CostoProveedorPresupuesto AS costo ON costo.CostoProveedorId = pago.CostoProveedorId
        WHERE pago.CostoProveedorId = @CostoProveedorId
        ORDER BY pago.FechaPago DESC, pago.PagoProveedorId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Tablas/CierrePersonalHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CierrePersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CierrePersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CierrePersonalHYT
    (
        Mes DATE NOT NULL CONSTRAINT PK_CierrePersonalHYT PRIMARY KEY,
        FechaCierre DATETIME2(3) NOT NULL CONSTRAINT DF_CierrePersonalHYT_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioId NVARCHAR(128) NOT NULL CONSTRAINT FK_CierrePersonalHYT_Usuario REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_CierrePersonalHYT_Mes CHECK (DAY(Mes) = 1)
    );
END;
GO

-- Fuente: Tablas/CeseTrabajadorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CeseTrabajadorHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.CeseTrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CeseTrabajadorHYT
    (
        CeseId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CeseTrabajadorHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_CeseTrabajadorHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        FechaCese DATE NOT NULL,
        UltimoDiaTrabajado DATE NOT NULL,
        Motivo NVARCHAR(300) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_CeseTrabajadorHYT_Operacion UNIQUE,
        CONSTRAINT CK_CeseTrabajadorHYT_Fechas CHECK (UltimoDiaTrabajado <= FechaCese),
        CONSTRAINT UQ_CeseTrabajadorHYT_TrabajadorFecha UNIQUE (TrabajadorId, FechaCese)
    );
END;
GO

-- Fuente: Tablas/MovimientoPersonalHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.MovimientoPersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.MovimientoPersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MovimientoPersonalHYT
    (
        MovimientoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_MovimientoPersonalHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        CeseId BIGINT NULL CONSTRAINT FK_MovimientoPersonalHYT_Cese REFERENCES dbo.CeseTrabajadorHYT(CeseId),
        TipoTrabajador VARCHAR(10) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Trabajador CHECK (TipoTrabajador IN ('PLANILLA', 'EXTERNO')),
        Concepto VARCHAR(24) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Concepto CHECK
            (Concepto IN ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES',
                          'ADELANTO', 'ADELANTO_GRATIFICACION', 'DESCUENTO', 'PAGO')),
        Fecha DATE NOT NULL,
        Descripcion NVARCHAR(300) NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_MovimientoPersonalHYT_Importe CHECK (Importe > 0),
        MonedaCodigo NVARCHAR(3) NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Moneda REFERENCES dbo.Moneda(Codigo),
        Frecuencia VARCHAR(10) NULL,
        PeriodoDesde DATE NULL,
        PeriodoHasta DATE NULL,
        Referencia NVARCHAR(100) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_MovimientoPersonalHYT_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_MovimientoPersonalHYT_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioId NVARCHAR(128) NOT NULL CONSTRAINT FK_MovimientoPersonalHYT_Usuario REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_MovimientoPersonalHYT_Periodo CHECK (
            (Frecuencia IS NULL AND PeriodoDesde IS NULL AND PeriodoHasta IS NULL)
            OR (Frecuencia IS NOT NULL AND Frecuencia IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
                AND PeriodoDesde IS NOT NULL AND PeriodoHasta IS NOT NULL AND PeriodoDesde <= PeriodoHasta))
    );
END;
GO

-- Fuente: Tablas/ObligacionPersonalHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.ObligacionPersonalHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.ObligacionPersonalHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ObligacionPersonalHYT
    (
        ObligacionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ObligacionPersonalHYT PRIMARY KEY,
        Tipo VARCHAR(3) NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Tipo CHECK (Tipo IN ('AFP', 'ONP')),
        Entidad NVARCHAR(120) NOT NULL,
        Mes DATE NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Mes CHECK (DAY(Mes) = 1),
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_ObligacionPersonalHYT_Importe CHECK (Importe > 0),
        Vencimiento DATE NOT NULL,
        FechaPago DATE NULL,
        ReferenciaPago NVARCHAR(100) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_ObligacionPersonalHYT_TipoEntidadMes UNIQUE (Tipo, Entidad, Mes),
        CONSTRAINT CK_ObligacionPersonalHYT_Pago CHECK (
            (FechaPago IS NULL AND ReferenciaPago IS NULL) OR
            (FechaPago IS NOT NULL AND ReferenciaPago IS NOT NULL AND LEN(ReferenciaPago) > 0))
    );
END;
GO

-- Fuente: Tablas/PolizaSctrHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PolizaSctrHYT para las etapas 3 y 4 de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
IF OBJECT_ID(N'dbo.PolizaSctrHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PolizaSctrHYT
    (
        PolizaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PolizaSctrHYT PRIMARY KEY,
        Aseguradora NVARCHAR(120) NOT NULL,
        NumeroPoliza NVARCHAR(80) NOT NULL,
        Cobertura VARCHAR(10) NOT NULL CONSTRAINT CK_PolizaSctrHYT_Cobertura CHECK (Cobertura IN ('SALUD', 'PENSION')),
        VigenciaDesde DATE NOT NULL,
        VigenciaHasta DATE NOT NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_PolizaSctrHYT_Fechas CHECK (VigenciaDesde <= VigenciaHasta),
        CONSTRAINT UQ_PolizaSctrHYT_Vigencia UNIQUE (Aseguradora, NumeroPoliza, Cobertura, VigenciaDesde)
    );
END;
GO

-- Fuente: StoreProcedure/spu_MovimientoPersonalHYT_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_MovimientoPersonalHYT_Registrar para pagos del personal.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MovimientoPersonalHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @TrabajadorId BIGINT,
    @Concepto VARCHAR(24),
    @Fecha DATE,
    @Descripcion NVARCHAR(300),
    @Importe DECIMAL(20,4),
    @MonedaCodigo NVARCHAR(3),
    @Frecuencia VARCHAR(10) = NULL,
    @PeriodoDesde DATE = NULL,
    @PeriodoHasta DATE = NULL,
    @Referencia NVARCHAR(100) = NULL,
    @CeseId BIGINT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar movimientos del personal.', 1;

        SET @Descripcion = NULLIF(TRIM(@Descripcion), N'');
        SET @Referencia = NULLIF(TRIM(@Referencia), N'');
        IF @OperacionId IS NULL OR @Fecha IS NULL OR @Descripcion IS NULL OR @Importe IS NULL
           OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2) OR @Concepto IS NULL
           OR @Concepto NOT IN ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION',
               'VACACIONES_PENDIENTES', 'ADELANTO', 'ADELANTO_GRATIFICACION', 'DESCUENTO', 'PAGO')
            THROW 51000, N'Revise el concepto, la fecha y el importe.', 1;

        IF (@Frecuencia IS NULL AND (@PeriodoDesde IS NOT NULL OR @PeriodoHasta IS NOT NULL))
           OR (@Frecuencia IS NOT NULL AND (@Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL')
               OR @PeriodoDesde IS NULL OR @PeriodoHasta IS NULL OR @PeriodoHasta < @PeriodoDesde))
            THROW 51000, N'Revise la frecuencia y las fechas de cobertura.', 1;

        IF @Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION', 'PAGO') AND @Referencia IS NULL
            THROW 51000, N'Indique la referencia del pago o adelanto.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        DECLARE @MovimientoId BIGINT;
        SELECT @MovimientoId = movimiento.MovimientoId
        FROM dbo.MovimientoPersonalHYT AS movimiento
        WHERE movimiento.OperacionId = @OperacionId;
        IF @MovimientoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.MovimientoPersonalHYT AS movimiento
                WHERE movimiento.MovimientoId = @MovimientoId AND movimiento.UsuarioId = @UsuarioId
                  AND movimiento.TrabajadorId = @TrabajadorId AND movimiento.Concepto = @Concepto
                  AND movimiento.Fecha = @Fecha AND movimiento.Descripcion = @Descripcion
                  AND movimiento.Importe = @Importe AND movimiento.MonedaCodigo = @MonedaCodigo
                  AND movimiento.Frecuencia IS NOT DISTINCT FROM @Frecuencia
                  AND movimiento.PeriodoDesde IS NOT DISTINCT FROM @PeriodoDesde
                  AND movimiento.PeriodoHasta IS NOT DISTINCT FROM @PeriodoHasta
                  AND movimiento.Referencia IS NOT DISTINCT FROM @Referencia
                  AND movimiento.CeseId IS NOT DISTINCT FROM @CeseId)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Mes DATE = DATEFROMPARTS(YEAR(@Fecha), MONTH(@Fecha), 1);
            IF EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes >= @Mes)
                THROW 51000, N'El movimiento corresponde a un mes cerrado o anterior al ultimo cierre.', 1;

            DECLARE @Tipo VARCHAR(10), @Ingreso DATE;
            SELECT @Tipo = trabajador.Tipo, @Ingreso = trabajador.FechaIngreso
            FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
            WHERE trabajador.TrabajadorId = @TrabajadorId;
            IF @Tipo IS NULL OR @Fecha < @Ingreso
                THROW 51000, N'El trabajador o la fecha del movimiento no son validos.', 1;
            IF (@Tipo = 'EXTERNO' AND @Concepto IN ('REMUNERACION', 'GRATIFICACION',
                'ADELANTO_GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES'))
                OR (@Tipo = 'PLANILLA' AND @Concepto = 'SERVICIO')
                THROW 51000, N'El concepto no corresponde al tipo de trabajador.', 1;

            IF (@Concepto IN ('LIQUIDACION', 'VACACIONES_PENDIENTES') AND @CeseId IS NULL)
                OR (@CeseId IS NOT NULL AND NOT EXISTS (
                    SELECT 1 FROM dbo.CeseTrabajadorHYT AS cese
                    WHERE cese.CeseId = @CeseId AND cese.TrabajadorId = @TrabajadorId AND cese.FechaCese <= @Fecha))
                THROW 51000, N'Seleccione un cese valido del trabajador.', 1;

            IF @Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION', 'PAGO', 'DESCUENTO')
                AND EXISTS (
                    SELECT 1 FROM dbo.MovimientoPersonalHYT AS posterior
                    WHERE posterior.TrabajadorId = @TrabajadorId AND posterior.MonedaCodigo = @MonedaCodigo
                      AND posterior.Fecha > @Fecha AND posterior.Concepto IN ('PAGO', 'DESCUENTO'))
                THROW 51000, N'Hay pagos o descuentos posteriores a la fecha del movimiento.', 1;

            IF @Concepto IN ('PAGO', 'DESCUENTO')
            BEGIN
                DECLARE @Saldo DECIMAL(38,2);
                SELECT @Saldo = COALESCE(SUM(CASE WHEN movimiento.Concepto IN
                    ('REMUNERACION', 'SERVICIO', 'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                    THEN movimiento.Importe ELSE -movimiento.Importe END), 0)
                FROM dbo.MovimientoPersonalHYT AS movimiento
                WHERE movimiento.TrabajadorId = @TrabajadorId AND movimiento.MonedaCodigo = @MonedaCodigo
                  AND movimiento.Fecha <= @Fecha;
                IF @Importe > @Saldo
                    THROW 51000, N'El importe excede el saldo del trabajador a esa fecha.', 1;
            END;

            INSERT INTO dbo.MovimientoPersonalHYT
                (TrabajadorId, CeseId, TipoTrabajador, Concepto, Fecha, Descripcion, Importe,
                 MonedaCodigo, Frecuencia, PeriodoDesde, PeriodoHasta, Referencia, OperacionId, UsuarioId)
            VALUES (@TrabajadorId, @CeseId, @Tipo, @Concepto, @Fecha, @Descripcion, @Importe,
                    @MonedaCodigo, @Frecuencia, @PeriodoDesde, @PeriodoHasta, @Referencia, @OperacionId, @UsuarioId);

            SET @MovimientoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @MovimientoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'MovimientoPersonalHYT',
                 @RegistroId, N'REGISTRAR', @EstadoNuevo = @Concepto;
        END;

        COMMIT TRANSACTION;
        SELECT @MovimientoId AS MovimientoId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_CierrePersonalHYT_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CierrePersonalHYT_Registrar para pagos del personal.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CierrePersonalHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CERRAR') = 0
            THROW 51000, N'No tiene permiso para cerrar el mes.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @Mes IS NULL OR DAY(@Mes) <> 1 OR EOMONTH(@Mes) >= @Hoy
            THROW 51000, N'Seleccione un mes que ya haya terminado.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        IF NOT EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes = @Mes)
        BEGIN
            IF EXISTS (
                SELECT 1 FROM dbo.MovimientoPersonalHYT AS movimiento
                LEFT JOIN dbo.CierrePersonalHYT AS cierre
                  ON cierre.Mes = DATEFROMPARTS(YEAR(movimiento.Fecha), MONTH(movimiento.Fecha), 1)
                WHERE movimiento.Fecha < @Mes AND cierre.Mes IS NULL)
                THROW 51000, N'Cierre primero los meses anteriores que tienen movimientos.', 1;

            INSERT INTO dbo.CierrePersonalHYT (Mes, UsuarioId) VALUES (@Mes, @UsuarioId);

            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(10), @Mes, 23);
            DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'CierrePersonalHYT',
                 @RegistroId, N'CERRAR', @EstadoNuevo = N'CERRADO';
        END;

        COMMIT TRANSACTION;
        SELECT @Mes AS Mes;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_EstadoCuentaPersonalHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_EstadoCuentaPersonalHYT_Consultar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_EstadoCuentaPersonalHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar pagos del personal.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hasta DATE = EOMONTH(@Mes);
        SELECT (
            SELECT trabajador.TrabajadorId, trabajador.NombreCompleto, trabajador.Tipo, trabajador.Activo,
                   @Mes AS Mes, CONVERT(BIT, CASE WHEN cierre.Mes IS NULL THEN 0 ELSE 1 END) AS Cerrado,
                   JSON_QUERY((
                       SELECT movimiento.MonedaCodigo,
                              SUM(CASE WHEN movimiento.Fecha < @Mes THEN signo.Neto ELSE 0 END) AS SaldoAnterior,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND signo.Neto > 0 THEN movimiento.Importe ELSE 0 END) AS Devengado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION') THEN movimiento.Importe ELSE 0 END) AS Adelantado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'DESCUENTO' THEN movimiento.Importe ELSE 0 END) AS Descontado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'PAGO' THEN movimiento.Importe ELSE 0 END) AS Pagado,
                              SUM(signo.Neto) AS Saldo
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       CROSS APPLY (VALUES (CASE WHEN movimiento.Concepto IN ('REMUNERACION', 'SERVICIO',
                           'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                           THEN movimiento.Importe ELSE -movimiento.Importe END)) AS signo (Neto)
                       WHERE movimiento.TrabajadorId = trabajador.TrabajadorId AND movimiento.Fecha <= @Hasta
                       GROUP BY movimiento.MonedaCodigo
                       ORDER BY movimiento.MonedaCodigo
                       FOR JSON PATH
                   )) AS Saldos,
                   JSON_QUERY((
                       SELECT movimiento.MovimientoId, movimiento.TipoTrabajador, movimiento.Concepto,
                              movimiento.Fecha, movimiento.Descripcion, movimiento.Importe, movimiento.MonedaCodigo,
                              movimiento.Frecuencia, movimiento.PeriodoDesde, movimiento.PeriodoHasta,
                              movimiento.Referencia, movimiento.CeseId
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       WHERE movimiento.TrabajadorId = trabajador.TrabajadorId AND movimiento.Fecha BETWEEN @Mes AND @Hasta
                       ORDER BY movimiento.Fecha DESC, movimiento.MovimientoId DESC
                       FOR JSON PATH
                   )) AS Movimientos,
                   JSON_QUERY((
                       SELECT cese.CeseId, cese.FechaCese, cese.UltimoDiaTrabajado, cese.Motivo
                       FROM dbo.CeseTrabajadorHYT AS cese
                       WHERE cese.TrabajadorId = trabajador.TrabajadorId
                       ORDER BY cese.FechaCese DESC, cese.CeseId DESC
                       FOR JSON PATH
                   )) AS Ceses
            FROM dbo.TrabajadorHYT AS trabajador
            LEFT JOIN dbo.CierrePersonalHYT AS cierre ON cierre.Mes = @Mes
            WHERE trabajador.TrabajadorId = @TrabajadorId
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_CeseTrabajadorHYT_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CeseTrabajadorHYT_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CeseTrabajadorHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @TrabajadorId BIGINT,
    @FechaCese DATE,
    @UltimoDiaTrabajado DATE,
    @Motivo NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CESES') = 0
            THROW 51000, N'No tiene permiso para registrar ceses.', 1;

        SET @Motivo = NULLIF(TRIM(@Motivo), N'');
        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @OperacionId IS NULL OR @FechaCese IS NULL OR @UltimoDiaTrabajado IS NULL
           OR @UltimoDiaTrabajado > @FechaCese OR @FechaCese > @Hoy
            THROW 51000, N'Revise la fecha de cese y el ultimo dia trabajado.', 1;

        BEGIN TRANSACTION;
        DECLARE @Bloqueo INT;
        EXEC @Bloqueo = sys.sp_getapplock @Resource = N'HYT.Planillas.Personal',
             @LockMode = 'Exclusive', @LockOwner = 'Transaction', @LockTimeout = 10000;
        IF @Bloqueo < 0
            THROW 51000, N'Planillas esta procesando otra operacion. Vuelva a intentar.', 1;

        DECLARE @CeseId BIGINT;
        SELECT @CeseId = cese.CeseId FROM dbo.CeseTrabajadorHYT AS cese WHERE cese.OperacionId = @OperacionId;
        IF @CeseId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.CeseTrabajadorHYT AS cese
                WHERE cese.CeseId = @CeseId AND cese.TrabajadorId = @TrabajadorId
                  AND cese.FechaCese = @FechaCese AND cese.UltimoDiaTrabajado = @UltimoDiaTrabajado
                  AND cese.Motivo IS NOT DISTINCT FROM @Motivo)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Mes DATE = DATEFROMPARTS(YEAR(@FechaCese), MONTH(@FechaCese), 1);
            IF EXISTS (SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes >= @Mes)
                THROW 51000, N'El cese corresponde a un mes cerrado o anterior al ultimo cierre.', 1;

            UPDATE dbo.TrabajadorHYT
            SET Activo = 0
            WHERE TrabajadorId = @TrabajadorId AND Activo = 1 AND FechaIngreso <= @UltimoDiaTrabajado;
            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador no esta activo o el ultimo dia es anterior al ingreso.', 1;

            INSERT INTO dbo.CeseTrabajadorHYT (TrabajadorId, FechaCese, UltimoDiaTrabajado, Motivo, OperacionId)
            VALUES (@TrabajadorId, @FechaCese, @UltimoDiaTrabajado, @Motivo, @OperacionId);

            SET @CeseId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CeseId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'CeseTrabajadorHYT',
                 @RegistroId, N'REGISTRAR', @EstadoNuevo = N'CESADO', @FechaEvento = @FechaCese, @Motivo = @Motivo;
        END;

        COMMIT TRANSACTION;
        SELECT @CeseId AS CeseId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ReportePlanillasHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Concilia el mes de personal y proveedores por moneda.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ReportePlanillasHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE,
    @ProveedorId BIGINT = NULL,
    @PresupuestoId BIGINT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'REPORTES_PLANILLAS.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar reportes de planillas.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hasta DATE = EOMONTH(@Mes);
        DECLARE @HastaUtc DATETIME2(0) = DATEADD(HOUR, 5, DATEADD(DAY, 1, CONVERT(DATETIME2(0), @Hasta)));
        DECLARE @VerPersonal BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PERSONAL.CONSULTAR');
        DECLARE @VerProveedores BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR');

        SELECT (
            SELECT @Mes AS Mes, @VerPersonal AS VerPersonal, @VerProveedores AS VerProveedores,
                   CONVERT(BIT, CASE WHEN @VerPersonal = 1 AND EXISTS (
                       SELECT 1 FROM dbo.CierrePersonalHYT AS cierre WHERE cierre.Mes = @Mes)
                       THEN 1 ELSE 0 END) AS Cerrado,
                   JSON_QUERY((
                       SELECT movimiento.TrabajadorId, trabajador.NombreCompleto,
                              movimiento.TipoTrabajador, movimiento.MonedaCodigo,
                              SUM(CASE WHEN movimiento.Fecha < @Mes THEN signo.Neto ELSE 0 END) AS SaldoAnterior,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND signo.Neto > 0 THEN movimiento.Importe ELSE 0 END) AS Devengado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto IN ('ADELANTO', 'ADELANTO_GRATIFICACION')
                                  THEN movimiento.Importe ELSE 0 END) AS Adelantado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'DESCUENTO' THEN movimiento.Importe ELSE 0 END) AS Descontado,
                              SUM(CASE WHEN movimiento.Fecha >= @Mes AND movimiento.Concepto = 'PAGO' THEN movimiento.Importe ELSE 0 END) AS Pagado,
                              SUM(signo.Neto) AS Saldo
                       FROM dbo.MovimientoPersonalHYT AS movimiento
                       INNER JOIN dbo.TrabajadorHYT AS trabajador ON trabajador.TrabajadorId = movimiento.TrabajadorId
                       CROSS APPLY (VALUES (CASE WHEN movimiento.Concepto IN ('REMUNERACION', 'SERVICIO',
                           'GRATIFICACION', 'LIQUIDACION', 'VACACIONES_PENDIENTES')
                           THEN movimiento.Importe ELSE -movimiento.Importe END)) AS signo (Neto)
                       WHERE @VerPersonal = 1 AND movimiento.Fecha <= @Hasta
                       GROUP BY movimiento.TrabajadorId, trabajador.NombreCompleto, movimiento.TipoTrabajador, movimiento.MonedaCodigo
                       HAVING MAX(movimiento.Fecha) >= @Mes OR SUM(signo.Neto) <> 0
                       ORDER BY movimiento.TipoTrabajador, trabajador.NombreCompleto, movimiento.TrabajadorId, movimiento.MonedaCodigo
                       FOR JSON PATH
                   )) AS Personal,
                   JSON_QUERY((
                       SELECT costo.ProveedorId, proveedor.RazonSocial, costo.PresupuestoId, presupuesto.NumeroInterno,
                              costo.MonedaCodigo, SUM(costo.Importe) AS Costo,
                              SUM(COALESCE(abonos.PagadoMes, 0)) AS PagadoMes,
                              SUM(COALESCE(abonos.PagadoAcumulado, 0)) AS PagadoAcumulado,
                              SUM(CONVERT(DECIMAL(19,2), costo.Importe - COALESCE(abonos.PagadoAcumulado, 0))) AS Pendiente
                       FROM dbo.CostoProveedorPresupuesto AS costo
                       INNER JOIN dbo.ProveedorHYT AS proveedor ON proveedor.ProveedorId = costo.ProveedorId
                       INNER JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId = costo.PresupuestoId
                       OUTER APPLY (
                           SELECT CONVERT(DECIMAL(18,2), SUM(CASE WHEN pago.FechaPago >= @Mes THEN pago.Importe ELSE 0 END)) AS PagadoMes,
                                  CONVERT(DECIMAL(18,2), SUM(pago.Importe)) AS PagadoAcumulado
                           FROM dbo.PagoProveedorHYT AS pago
                           WHERE pago.CostoProveedorId = costo.CostoProveedorId AND pago.FechaPago <= @Hasta
                       ) AS abonos
                       WHERE @VerProveedores = 1
                         AND (costo.FechaRegistro < @HastaUtc OR abonos.PagadoAcumulado IS NOT NULL)
                         AND (@ProveedorId IS NULL OR costo.ProveedorId = @ProveedorId)
                         AND (@PresupuestoId IS NULL OR costo.PresupuestoId = @PresupuestoId)
                       GROUP BY costo.ProveedorId, proveedor.RazonSocial, costo.PresupuestoId, presupuesto.NumeroInterno, costo.MonedaCodigo
                       HAVING SUM(COALESCE(abonos.PagadoMes, 0)) > 0 OR SUM(costo.Importe) > SUM(COALESCE(abonos.PagadoAcumulado, 0))
                       ORDER BY proveedor.RazonSocial, presupuesto.NumeroInterno, costo.MonedaCodigo
                       FOR JSON PATH
                   )) AS Proveedores
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/035_PermisosPagosPersonal.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra permisos de estado de cuenta, movimientos, cierres, ceses y reportes de planillas.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'PAGOS_PERSONAL.CONSULTAR', N'Consultar el estado de cuenta del personal'),
    (N'PAGOS_PERSONAL.EDITAR', N'Registrar movimientos de personal'),
    (N'PAGOS_PERSONAL.CERRAR', N'Cerrar el mes de movimientos de personal'),
    (N'PAGOS_PERSONAL.CESES', N'Registrar ceses de personal'),
    (N'REPORTES_PLANILLAS.CONSULTAR', N'Consultar y exportar reportes de planillas')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_MovimientoPersonalHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CierrePersonalHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_EstadoCuentaPersonalHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CeseTrabajadorHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ReportePlanillasHYT_Consultar TO hyt_app_executor;
GO

-- Fuente: StoreProcedure/spu_ObligacionesHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ObligacionesHYT_Consultar para obligaciones de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObligacionesHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @Mes DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar obligaciones.', 1;
        IF @Mes IS NULL OR DAY(@Mes) <> 1
            THROW 51000, N'Seleccione un mes valido.', 1;

        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        SELECT (
            SELECT @Mes AS Mes,
                   JSON_QUERY((
                       SELECT obligacion.ObligacionId, obligacion.Tipo, obligacion.Entidad, obligacion.Mes,
                              obligacion.Importe, obligacion.Vencimiento, obligacion.FechaPago,
                              obligacion.ReferenciaPago, obligacion.VersionFila,
                              CASE WHEN obligacion.FechaPago IS NOT NULL THEN 'PAGADA'
                                   WHEN obligacion.Vencimiento < @Hoy THEN 'VENCIDA' ELSE 'PENDIENTE' END AS Estado
                       FROM dbo.ObligacionPersonalHYT AS obligacion
                       WHERE obligacion.Mes = @Mes
                       ORDER BY obligacion.Vencimiento, obligacion.Tipo, obligacion.Entidad
                       FOR JSON PATH
                   )) AS Obligaciones,
                   JSON_QUERY((
                       SELECT poliza.PolizaId, poliza.Aseguradora, poliza.NumeroPoliza, poliza.Cobertura,
                              poliza.VigenciaDesde, poliza.VigenciaHasta, poliza.VersionFila,
                              CASE WHEN poliza.VigenciaHasta < @Hoy THEN 'VENCIDA'
                                   WHEN poliza.VigenciaDesde > @Hoy THEN 'PROGRAMADA'
                                   WHEN DATEDIFF(DAY, @Hoy, poliza.VigenciaHasta) <= 30 THEN 'POR_VENCER'
                                   ELSE 'VIGENTE' END AS Estado
                       FROM dbo.PolizaSctrHYT AS poliza
                       ORDER BY poliza.VigenciaHasta DESC, poliza.PolizaId DESC
                       FOR JSON PATH
                   )) AS Polizas
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        ) AS Datos;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ObligacionPersonalHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ObligacionPersonalHYT_Guardar para obligaciones de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObligacionPersonalHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Tipo VARCHAR(3),
    @Entidad NVARCHAR(120),
    @Mes DATE,
    @Importe DECIMAL(20,4),
    @Vencimiento DATE,
    @FechaPago DATE = NULL,
    @ReferenciaPago NVARCHAR(100) = NULL,
    @ObligacionId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar obligaciones.', 1;

        SET @Entidad = NULLIF(TRIM(@Entidad), N'');
        SET @ReferenciaPago = NULLIF(TRIM(@ReferenciaPago), N'');
        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @Tipo IS NULL OR @Tipo NOT IN ('AFP', 'ONP') OR @Entidad IS NULL OR @Mes IS NULL
           OR DAY(@Mes) <> 1 OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
           OR @Vencimiento IS NULL
           OR (@FechaPago IS NULL AND @ReferenciaPago IS NOT NULL)
           OR (@FechaPago IS NOT NULL AND (@ReferenciaPago IS NULL OR @FechaPago > @Hoy))
            THROW 51000, N'Revise la entidad, el mes, el importe y las fechas de la obligacion.', 1;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @ObligacionId IS NULL
        BEGIN
            INSERT INTO dbo.ObligacionPersonalHYT (Tipo, Entidad, Mes, Importe, Vencimiento, FechaPago, ReferenciaPago)
            VALUES (@Tipo, @Entidad, @Mes, @Importe, @Vencimiento, @FechaPago, @ReferenciaPago);
            SET @ObligacionId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.ObligacionPersonalHYT
            SET Tipo = @Tipo, Entidad = @Entidad, Mes = @Mes, Importe = @Importe, Vencimiento = @Vencimiento,
                FechaPago = @FechaPago, ReferenciaPago = @ReferenciaPago
            WHERE ObligacionId = @ObligacionId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La obligacion cambio o no existe. Vuelva a abrirla.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @ObligacionId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Tipo AS Tipo, @Entidad AS Entidad, @Mes AS Mes, @Importe AS Importe,
                   @Vencimiento AS Vencimiento, @FechaPago AS FechaPago, @ReferenciaPago AS ReferenciaPago
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'ObligacionPersonalHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @ObligacionId AS ObligacionId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PolizaSctrHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_PolizaSctrHYT_Guardar para obligaciones de Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PolizaSctrHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Aseguradora NVARCHAR(120),
    @NumeroPoliza NVARCHAR(80),
    @Cobertura VARCHAR(10),
    @VigenciaDesde DATE,
    @VigenciaHasta DATE,
    @PolizaId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar polizas SCTR.', 1;

        SET @Aseguradora = NULLIF(TRIM(@Aseguradora), N'');
        SET @NumeroPoliza = NULLIF(TRIM(@NumeroPoliza), N'');
        IF @Aseguradora IS NULL OR @NumeroPoliza IS NULL OR @Cobertura IS NULL
           OR @Cobertura NOT IN ('SALUD', 'PENSION') OR @VigenciaDesde IS NULL
           OR @VigenciaHasta IS NULL OR @VigenciaHasta < @VigenciaDesde
            THROW 51000, N'Revise la aseguradora, la cobertura y la vigencia de la poliza.', 1;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @PolizaId IS NULL
        BEGIN
            INSERT INTO dbo.PolizaSctrHYT (Aseguradora, NumeroPoliza, Cobertura, VigenciaDesde, VigenciaHasta)
            VALUES (@Aseguradora, @NumeroPoliza, @Cobertura, @VigenciaDesde, @VigenciaHasta);
            SET @PolizaId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.PolizaSctrHYT
            SET Aseguradora = @Aseguradora, NumeroPoliza = @NumeroPoliza, Cobertura = @Cobertura,
                VigenciaDesde = @VigenciaDesde, VigenciaHasta = @VigenciaHasta
            WHERE PolizaId = @PolizaId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La poliza cambio o no existe. Vuelva a abrirla.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PolizaId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Aseguradora AS Aseguradora, @NumeroPoliza AS NumeroPoliza, @Cobertura AS Cobertura,
                   @VigenciaDesde AS VigenciaDesde, @VigenciaHasta AS VigenciaHasta
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PolizaSctrHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @PolizaId AS PolizaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/036_PermisosObligaciones.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Autoriza consulta y edicion de obligaciones y polizas SCTR.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'OBLIGACIONES.CONSULTAR', N'Consultar AFP, ONP y polizas SCTR'),
    (N'OBLIGACIONES.EDITAR', N'Guardar AFP, ONP y polizas SCTR')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ObligacionesHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ObligacionPersonalHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PolizaSctrHYT_Guardar TO hyt_app_executor;
GO

COMMIT TRANSACTION;
GO
