-- Firma: FRANCO LARA | Create date: 09/11/2026 | Proveedores, costos, abonos, trabajadores, vacaciones y asistencia.
-- Paquete incremental. Ejecutar con sqlcmd -b -f 65001 después de Administración.
USE [SistemaWebHYT];
GO
SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
SET XACT_ABORT ON;
BEGIN TRANSACTION;
GO

-- Fuente: Tablas/ProveedorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.ProveedorHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ProveedorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ProveedorHYT
    (
        ProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ProveedorHYT PRIMARY KEY,
        Ruc VARCHAR(11) NOT NULL CONSTRAINT UQ_ProveedorHYT_Ruc UNIQUE,
        RazonSocial NVARCHAR(180) NOT NULL,
        Banco NVARCHAR(80) NULL,
        NumeroCuenta NVARCHAR(34) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_ProveedorHYT_Activo DEFAULT (1),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_ProveedorHYT_Ruc CHECK (LEN(Ruc) = 11 AND Ruc NOT LIKE '%[^0-9]%'),
        CONSTRAINT CK_ProveedorHYT_Cuenta CHECK ((Banco IS NULL AND NumeroCuenta IS NULL)
            OR (Banco IS NOT NULL AND NumeroCuenta IS NOT NULL AND LEN(Banco) > 0 AND LEN(NumeroCuenta) > 0))
    );
END;
GO

-- Fuente: Tablas/CostoProveedorPresupuesto.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.CostoProveedorPresupuesto para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.CostoProveedorPresupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CostoProveedorPresupuesto
    (
        CostoProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_CostoProveedorPresupuesto PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL CONSTRAINT FK_CostoProveedor_Presupuesto REFERENCES dbo.Presupuesto(PresupuestoId),
        ProveedorId BIGINT NOT NULL CONSTRAINT FK_CostoProveedor_Proveedor REFERENCES dbo.ProveedorHYT(ProveedorId),
        Descripcion NVARCHAR(300) NOT NULL,
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_CostoProveedor_Importe CHECK (Importe > 0),
        MonedaCodigo NVARCHAR(3) NOT NULL CONSTRAINT FK_CostoProveedor_Moneda REFERENCES dbo.Moneda(Codigo),
        DocumentoSustento NVARCHAR(100) NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_CostoProveedor_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_CostoProveedor_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL CONSTRAINT FK_CostoProveedor_Usuario REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO

-- Fuente: Tablas/PagoProveedorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PagoProveedorHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PagoProveedorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PagoProveedorHYT
    (
        PagoProveedorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PagoProveedorHYT PRIMARY KEY,
        CostoProveedorId BIGINT NOT NULL CONSTRAINT FK_PagoProveedor_Costo REFERENCES dbo.CostoProveedorPresupuesto(CostoProveedorId),
        Importe DECIMAL(18,2) NOT NULL CONSTRAINT CK_PagoProveedor_Importe CHECK (Importe > 0),
        FechaPago DATE NOT NULL,
        Referencia NVARCHAR(100) NOT NULL,
        OperacionId UNIQUEIDENTIFIER NOT NULL CONSTRAINT UQ_PagoProveedor_Operacion UNIQUE,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_PagoProveedor_Fecha DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL CONSTRAINT FK_PagoProveedor_Usuario REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO

-- Fuente: Tablas/TrabajadorHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.TrabajadorHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TrabajadorHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TrabajadorHYT
    (
        TrabajadorId BIGINT IDENTITY NOT NULL CONSTRAINT PK_TrabajadorHYT PRIMARY KEY,
        Documento NVARCHAR(20) NOT NULL CONSTRAINT UQ_TrabajadorHYT_Documento UNIQUE,
        NombreCompleto NVARCHAR(180) NOT NULL,
        Tipo VARCHAR(10) NOT NULL CONSTRAINT CK_TrabajadorHYT_Tipo CHECK (Tipo IN ('PLANILLA', 'EXTERNO')),
        FechaIngreso DATE NOT NULL,
        Banco NVARCHAR(80) NULL,
        NumeroCuenta NVARCHAR(34) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_TrabajadorHYT_Activo DEFAULT (1),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_TrabajadorHYT_Cuenta CHECK ((Banco IS NULL AND NumeroCuenta IS NULL)
            OR (Banco IS NOT NULL AND NumeroCuenta IS NOT NULL AND LEN(Banco) > 0 AND LEN(NumeroCuenta) > 0))
    );
END;
GO

-- Fuente: Tablas/PeriodoVacacionesHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.PeriodoVacacionesHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PeriodoVacacionesHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PeriodoVacacionesHYT
    (
        PeriodoVacacionesId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PeriodoVacacionesHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_PeriodoVacaciones_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        FechaDesde DATE NOT NULL,
        FechaHasta DATE NOT NULL,
        DiasReconocidos SMALLINT NOT NULL CONSTRAINT CK_PeriodoVacaciones_Dias CHECK (DiasReconocidos > 0),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_PeriodoVacaciones_Fechas CHECK (FechaHasta >= FechaDesde),
        CONSTRAINT UQ_PeriodoVacaciones_TrabajadorDesde UNIQUE (TrabajadorId, FechaDesde)
    );
END;
GO

-- Fuente: Tablas/VacacionesHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.VacacionesHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.VacacionesHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VacacionesHYT
    (
        VacacionesId BIGINT IDENTITY NOT NULL CONSTRAINT PK_VacacionesHYT PRIMARY KEY,
        PeriodoVacacionesId BIGINT NOT NULL CONSTRAINT FK_Vacaciones_Periodo REFERENCES dbo.PeriodoVacacionesHYT(PeriodoVacacionesId),
        FechaInicio DATE NOT NULL,
        FechaFin DATE NOT NULL,
        Dias SMALLINT NOT NULL,
        Estado VARCHAR(10) NOT NULL CONSTRAINT CK_Vacaciones_Estado CHECK (Estado IN ('PENDIENTE', 'APROBADA', 'ANULADA')),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT CK_Vacaciones_FechasDias CHECK (FechaFin >= FechaInicio AND Dias BETWEEN 1 AND DATEDIFF(DAY, FechaInicio, FechaFin) + 1)
    );
END;
GO

-- Fuente: Tablas/JornadaHYT.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Define dbo.JornadaHYT para el modulo Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.JornadaHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.JornadaHYT
    (
        JornadaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_JornadaHYT PRIMARY KEY,
        TrabajadorId BIGINT NOT NULL CONSTRAINT FK_Jornada_Trabajador REFERENCES dbo.TrabajadorHYT(TrabajadorId),
        Entrada DATETIME2(0) NOT NULL,
        Salida DATETIME2(0) NULL,
        MinutosExtrasAprobados INT NOT NULL CONSTRAINT DF_Jornada_Extras DEFAULT (0),
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT UQ_Jornada_TrabajadorEntrada UNIQUE (TrabajadorId, Entrada),
        CONSTRAINT CK_Jornada_Horario CHECK (MinutosExtrasAprobados >= 0 AND
            ((Salida IS NULL AND MinutosExtrasAprobados = 0) OR
             (Salida IS NOT NULL AND Salida > Entrada AND MinutosExtrasAprobados <= DATEDIFF(MINUTE, Entrada, Salida))))
    );
END;
GO

-- Fuente: StoreProcedure/spu_CostoProveedor_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CostoProveedor_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CostoProveedor_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @PresupuestoId BIGINT,
    @ProveedorId BIGINT,
    @Descripcion NVARCHAR(300),
    @Importe DECIMAL(20,4),
    @MonedaCodigo NVARCHAR(3),
    @DocumentoSustento NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'COSTOS_PROVEEDOR.EDITAR') = 0
           OR dbo.fn_HYTPermiso(@UsuarioId, N'PRESUPUESTOS.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para registrar costos del presupuesto.', 1;

        SET @Descripcion = TRIM(@Descripcion);
        SET @DocumentoSustento = NULLIF(TRIM(@DocumentoSustento), N'');
        IF @OperacionId IS NULL OR NULLIF(@Descripcion, N'') IS NULL
           OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
            THROW 51000, N'Indique una descripcion y un importe positivo con hasta dos decimales.', 1;

        BEGIN TRANSACTION;

        DECLARE @CostoId BIGINT;
        SELECT @CostoId = costo.CostoProveedorId
        FROM dbo.CostoProveedorPresupuesto AS costo WITH (UPDLOCK, HOLDLOCK)
        WHERE costo.OperacionId = @OperacionId;

        IF @CostoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.CostoProveedorPresupuesto AS costo
                WHERE costo.CostoProveedorId = @CostoId AND costo.UsuarioRegistro = @UsuarioId
                    AND costo.PresupuestoId = @PresupuestoId AND costo.ProveedorId = @ProveedorId
                    AND costo.Descripcion = @Descripcion AND costo.Importe = @Importe
                    AND costo.MonedaCodigo = @MonedaCodigo
                    AND costo.DocumentoSustento IS NOT DISTINCT FROM @DocumentoSustento)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.ProveedorHYT AS proveedor WITH (HOLDLOCK)
                           WHERE proveedor.ProveedorId = @ProveedorId AND proveedor.Activo = 1)
                THROW 51000, N'Seleccione un proveedor activo.', 1;

            INSERT INTO dbo.CostoProveedorPresupuesto
                (PresupuestoId, ProveedorId, Descripcion, Importe, MonedaCodigo, DocumentoSustento, OperacionId, UsuarioRegistro)
            VALUES (@PresupuestoId, @ProveedorId, @Descripcion, @Importe, @MonedaCodigo,
                    @DocumentoSustento, @OperacionId, @UsuarioId);

            SET @CostoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CostoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'CostoProveedorPresupuesto', @RegistroId, N'REGISTRAR';
        END;

        COMMIT TRANSACTION;
        SELECT @CostoId AS CostoProveedorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PagoProveedor_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_PagoProveedor_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PagoProveedor_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @CostoProveedorId BIGINT,
    @Importe DECIMAL(20,4),
    @FechaPago DATE,
    @Referencia NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar pagos a proveedores.', 1;

        SET @Referencia = TRIM(@Referencia);
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
                    AND pago.FechaPago = @FechaPago AND pago.Referencia = @Referencia)
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
                (CostoProveedorId, Importe, FechaPago, Referencia, OperacionId, UsuarioRegistro)
            VALUES (@CostoProveedorId, @Importe, @FechaPago, @Referencia, @OperacionId, @UsuarioId);

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

-- Fuente: StoreProcedure/spu_CostoProveedor_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta costos, saldos y datos de pago segun los permisos de Planillas.
-- =============================================
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Agrega el comprobante de compra vinculado cuando está disponible el módulo de compras.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CostoProveedor_Listar
    @UsuarioId NVARCHAR(128),
    @PresupuestoId BIGINT = NULL,
    @ProveedorId BIGINT = NULL,
    @SoloPendientes BIT = 0,
    @Pagina INT = 1,
    @CostoProveedorId BIGINT = NULL,
    @Desde DATE = NULL,
    @Hasta DATE = NULL,
    @EstadoFiltro VARCHAR(10) = NULL,
    @Buscar NVARCHAR(180) = NULL
WITH EXECUTE AS OWNER
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR') = 0
           AND (@PresupuestoId IS NULL
                OR dbo.fn_HYTPermiso(@UsuarioId, N'COSTOS_PROVEEDOR.CONSULTAR') = 0
                OR dbo.fn_HYTPermiso(@UsuarioId, N'PRESUPUESTOS.CONSULTAR') = 0)
            THROW 51000, N'No tiene permiso para consultar costos de proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;
        IF @Desde > @Hasta OR @Hasta = '99991231'
           OR (@EstadoFiltro IS NOT NULL AND @EstadoFiltro NOT IN ('PENDIENTES', 'PAGADOS', 'TODOS'))
            THROW 51000, N'Revise el rango de fechas y el estado.', 1;

        DECLARE @DesdeUtc DATETIME2(0) = DATEADD(HOUR, 5, CONVERT(DATETIME2(0), @Desde));
        DECLARE @HastaUtc DATETIME2(0) = DATEADD(HOUR, 5, DATEADD(DAY, 1, CONVERT(DATETIME2(0), @Hasta)));
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA');
        DECLARE @VerPago BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.CONSULTAR');

        SELECT costo.CostoProveedorId, costo.PresupuestoId,
               COALESCE(presupuesto.NumeroPresupuesto, presupuesto.NumeroInterno) AS NumeroInterno,
               CONVERT(DATE, DATEADD(HOUR, -5, costo.FechaRegistro)) AS FechaCosto,
               costo.ProveedorId, proveedor.RazonSocial, costo.Descripcion, costo.MonedaCodigo,
               CASE WHEN @CostoProveedorId IS NOT NULL AND @VerPago = 1
                    THEN proveedor.Banco END AS Banco,
               CASE WHEN @CostoProveedorId IS NOT NULL AND @VerPago = 1 AND @VerCuenta = 1
                    THEN proveedor.NumeroCuenta END AS NumeroCuenta,
               costo.DocumentoSustento, costo.Importe, saldo.Pagado,
               costo.Importe - saldo.Pagado AS Pendiente,
               CASE WHEN saldo.Pagado = 0 THEN N'PENDIENTE'
                    WHEN saldo.Pagado = costo.Importe THEN N'PAGADO' ELSE N'PARCIAL' END AS Estado,
               COUNT_BIG(1) OVER () AS Total,
               CAST(NULL AS NVARCHAR(70)) AS FacturaCompra
        INTO #Pagina
        FROM dbo.CostoProveedorPresupuesto AS costo
        INNER JOIN dbo.Presupuesto AS presupuesto ON presupuesto.PresupuestoId = costo.PresupuestoId
        INNER JOIN dbo.ProveedorHYT AS proveedor ON proveedor.ProveedorId = costo.ProveedorId
        OUTER APPLY (
            SELECT CONVERT(DECIMAL(18,2), COALESCE(SUM(pago.Importe), 0)) AS Pagado
            FROM dbo.PagoProveedorHYT AS pago
            WHERE pago.CostoProveedorId = costo.CostoProveedorId
        ) AS saldo
        WHERE (@PresupuestoId IS NULL OR costo.PresupuestoId = @PresupuestoId)
          AND (@ProveedorId IS NULL OR costo.ProveedorId = @ProveedorId)
          AND (@CostoProveedorId IS NULL OR costo.CostoProveedorId = @CostoProveedorId)
          AND (@DesdeUtc IS NULL OR costo.FechaRegistro >= @DesdeUtc)
          AND (@HastaUtc IS NULL OR costo.FechaRegistro < @HastaUtc)
          AND (@Buscar IS NULL
               OR CHARINDEX(@Buscar, proveedor.RazonSocial COLLATE Latin1_General_100_CI_AI) > 0
               OR CHARINDEX(@Buscar, proveedor.Ruc) > 0
               OR CHARINDEX(@Buscar, costo.Descripcion COLLATE Latin1_General_100_CI_AI) > 0
               OR CHARINDEX(@Buscar, presupuesto.NumeroPresupuesto) > 0
               OR CHARINDEX(@Buscar, presupuesto.NumeroInterno) > 0)
          AND (@EstadoFiltro = 'TODOS'
               OR (@EstadoFiltro = 'PAGADOS' AND saldo.Pagado = costo.Importe)
               OR (@EstadoFiltro = 'PENDIENTES' AND saldo.Pagado < costo.Importe)
               OR (@EstadoFiltro IS NULL AND (@SoloPendientes = 0 OR saldo.Pagado < costo.Importe)))
        ORDER BY costo.CostoProveedorId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;

        IF @VerPago=1 AND OBJECT_ID(N'dbo.ComprobanteHYTCosto',N'U') IS NOT NULL
            AND OBJECT_ID(N'dbo.ComprobanteHYT',N'U') IS NOT NULL
        BEGIN
            UPDATE pagina SET FacturaCompra=CONCAT(compra.Serie,N'-',compra.Numero)
            FROM #Pagina AS pagina
            INNER JOIN dbo.ComprobanteHYTCosto AS vinculo ON vinculo.CostoProveedorId=pagina.CostoProveedorId
            INNER JOIN dbo.ComprobanteHYT AS compra ON compra.ComprobanteId=vinculo.ComprobanteId
            WHERE compra.EsVenta=0 AND compra.Serie IS NOT NULL AND compra.Numero IS NOT NULL;
        END;

        SELECT pagina.CostoProveedorId,pagina.PresupuestoId,pagina.NumeroInterno,pagina.FechaCosto,
            pagina.ProveedorId,pagina.RazonSocial,pagina.Descripcion,pagina.MonedaCodigo,pagina.Banco,pagina.NumeroCuenta,
            pagina.DocumentoSustento,pagina.Importe,pagina.Pagado,pagina.Pendiente,pagina.Estado,pagina.Total,pagina.FacturaCompra
        FROM #Pagina AS pagina ORDER BY pagina.CostoProveedorId DESC;
        DROP TABLE #Pagina;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Pagina',N'U') IS NOT NULL DROP TABLE #Pagina;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ProveedorHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ProveedorHYT_Listar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ProveedorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @ProveedorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT proveedor.ProveedorId, proveedor.Ruc, proveedor.RazonSocial, proveedor.Banco,
               CASE WHEN proveedor.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(proveedor.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @ProveedorId IS NOT NULL AND @VerCuenta = 1
                    THEN proveedor.NumeroCuenta END AS NumeroCuenta,
               proveedor.Activo, proveedor.VersionFila, COUNT_BIG(1) OVER () AS Total
        FROM dbo.ProveedorHYT AS proveedor
        WHERE (@ProveedorId IS NULL OR proveedor.ProveedorId = @ProveedorId)
          AND (@SoloActivos = 0 OR proveedor.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, proveedor.RazonSocial) > 0
               OR CHARINDEX(@Buscar, proveedor.Ruc) > 0)
        ORDER BY proveedor.RazonSocial, proveedor.ProveedorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_ProveedorHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ProveedorHYT_Guardar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ProveedorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Ruc VARCHAR(11),
    @RazonSocial NVARCHAR(180),
    @Activo BIT,
    @ProveedorId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ActualizarCuenta BIT = 0,
    @Banco NVARCHAR(80) = NULL,
    @NumeroCuenta NVARCHAR(34) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar proveedores.', 1;

        SET @RazonSocial = NULLIF(TRIM(@RazonSocial), N'');
        IF @Ruc IS NULL OR LEN(@Ruc) <> 11 OR @Ruc LIKE '%[^0-9]%'
           OR @RazonSocial IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique un RUC de once digitos y la razon social.', 1;

        SET @Banco = NULLIF(TRIM(@Banco), N'');
        SET @NumeroCuenta = NULLIF(TRIM(@NumeroCuenta), N'');
        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;
        END;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @ProveedorId IS NULL
        BEGIN
            INSERT INTO dbo.ProveedorHYT (Ruc, RazonSocial, Banco, NumeroCuenta, Activo)
            VALUES (@Ruc, @RazonSocial, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END, @Activo);
            SET @ProveedorId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.ProveedorHYT
            SET Ruc = @Ruc, RazonSocial = @RazonSocial, Activo = @Activo,
                Banco = CASE WHEN @ActualizarCuenta = 1 THEN @Banco ELSE Banco END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE ProveedorId = @ProveedorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El proveedor cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @ProveedorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Ruc AS Ruc, @RazonSocial AS RazonSocial, @Activo AS Activo,
                   @ActualizarCuenta AS ActualizoCuenta
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'ProveedorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;

        SELECT @ProveedorId AS ProveedorId;
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
-- Description:   Implementa spu_PagoProveedor_Listar para Planillas.
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

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_TrabajadorHYT_Listar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Tipo VARCHAR(10) = NULL,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar trabajadores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT trabajador.TrabajadorId, trabajador.Documento, trabajador.NombreCompleto, trabajador.Tipo,
               trabajador.FechaIngreso, trabajador.Banco,
               CASE WHEN trabajador.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(trabajador.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @TrabajadorId IS NOT NULL AND @VerCuenta = 1
                    THEN trabajador.NumeroCuenta END AS NumeroCuenta,
               trabajador.Activo, trabajador.VersionFila, COUNT_BIG(1) OVER () AS Total
        FROM dbo.TrabajadorHYT AS trabajador
        WHERE (@TrabajadorId IS NULL OR trabajador.TrabajadorId = @TrabajadorId)
          AND (@Tipo IS NULL OR trabajador.Tipo = @Tipo)
          AND (@SoloActivos = 0 OR trabajador.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, trabajador.NombreCompleto) > 0
               OR CHARINDEX(@Buscar, trabajador.Documento) > 0)
        ORDER BY trabajador.NombreCompleto, trabajador.TrabajadorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_TrabajadorHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_TrabajadorHYT_Guardar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_TrabajadorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Documento NVARCHAR(20),
    @NombreCompleto NVARCHAR(180),
    @Tipo VARCHAR(10),
    @FechaIngreso DATE,
    @Activo BIT,
    @TrabajadorId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ActualizarCuenta BIT = 0,
    @Banco NVARCHAR(80) = NULL,
    @NumeroCuenta NVARCHAR(34) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar trabajadores.', 1;

        SET @NombreCompleto = NULLIF(TRIM(@NombreCompleto), N'');
        SET @Documento = NULLIF(TRIM(@Documento), N'');
        IF @Documento IS NULL OR @Tipo IS NULL OR @Tipo NOT IN ('PLANILLA', 'EXTERNO')
           OR @FechaIngreso IS NULL
           OR @NombreCompleto IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique documento, nombre, tipo y fecha de ingreso.', 1;

        SET @Banco = NULLIF(TRIM(@Banco), N'');
        SET @NumeroCuenta = NULLIF(TRIM(@NumeroCuenta), N'');
        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'TRABAJADORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;
        END;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @TrabajadorId IS NULL
        BEGIN
            INSERT INTO dbo.TrabajadorHYT (Documento, NombreCompleto, Tipo, FechaIngreso, Banco, NumeroCuenta, Activo)
            VALUES (@Documento, @NombreCompleto, @Tipo, @FechaIngreso, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END, @Activo);
            SET @TrabajadorId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.TrabajadorHYT
            SET Documento = @Documento, NombreCompleto = @NombreCompleto, Activo = @Activo,
                Tipo = @Tipo, FechaIngreso = @FechaIngreso,
                Banco = CASE WHEN @ActualizarCuenta = 1 THEN @Banco ELSE Banco END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El trabajador cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @TrabajadorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Documento AS Documento, @NombreCompleto AS NombreCompleto, @Activo AS Activo,
                   @Tipo AS Tipo, @FechaIngreso AS FechaIngreso, @ActualizarCuenta AS ActualizoCuenta
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'TrabajadorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;

        SELECT @TrabajadorId AS TrabajadorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_VacacionesHYT_Consultar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_VacacionesHYT_Consultar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_VacacionesHYT_Consultar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Hoy DATE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar vacaciones.', 1;

        SELECT periodo.PeriodoVacacionesId, periodo.FechaDesde, periodo.FechaHasta,
               periodo.DiasReconocidos, periodo.VersionFila,
               COALESCE(saldo.Gozados, 0) AS DiasGozados,
               COALESCE(saldo.Programados, 0) AS DiasProgramados,
               periodo.DiasReconocidos - COALESCE(saldo.Gozados, 0) - COALESCE(saldo.Programados, 0) AS DiasDisponibles
        FROM dbo.PeriodoVacacionesHYT AS periodo
        OUTER APPLY (
            SELECT SUM(CASE WHEN vacacion.FechaFin < @Hoy THEN CONVERT(INT, vacacion.Dias) ELSE 0 END) AS Gozados,
                   SUM(CASE WHEN vacacion.FechaFin >= @Hoy THEN CONVERT(INT, vacacion.Dias) ELSE 0 END) AS Programados
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.PeriodoVacacionesId = periodo.PeriodoVacacionesId AND vacacion.Estado = 'APROBADA'
        ) AS saldo
        WHERE periodo.TrabajadorId = @TrabajadorId
        ORDER BY periodo.FechaDesde DESC;

        SELECT vacacion.VacacionesId, vacacion.PeriodoVacacionesId, vacacion.FechaInicio,
               vacacion.FechaFin, vacacion.Dias, vacacion.Estado, vacacion.VersionFila
        FROM dbo.VacacionesHYT AS vacacion
        INNER JOIN dbo.PeriodoVacacionesHYT AS periodo ON periodo.PeriodoVacacionesId = vacacion.PeriodoVacacionesId
        WHERE periodo.TrabajadorId = @TrabajadorId
        ORDER BY vacacion.FechaInicio DESC, vacacion.VacacionesId DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_PeriodoVacacionesHYT_Registrar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_PeriodoVacacionesHYT_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PeriodoVacacionesHYT_Registrar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @DiasReconocidos SMALLINT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar periodos de vacaciones.', 1;

        IF @FechaDesde IS NULL OR @FechaHasta IS NULL OR @FechaHasta < @FechaDesde
           OR @DiasReconocidos IS NULL OR @DiasReconocidos <= 0
            THROW 51000, N'Indique las fechas del periodo y los dias reconocidos.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.Tipo = 'PLANILLA';

        IF @FechaIngreso IS NULL OR @FechaDesde < @FechaIngreso
            THROW 51000, N'El periodo debe corresponder a un trabajador en planilla y ser posterior a su ingreso o comenzar en esa fecha.', 1;

        IF EXISTS (
            SELECT 1 FROM dbo.PeriodoVacacionesHYT AS periodo
            WHERE periodo.TrabajadorId = @TrabajadorId
              AND periodo.FechaDesde <= @FechaHasta AND periodo.FechaHasta >= @FechaDesde)
            THROW 51000, N'El periodo se superpone con otro periodo registrado.', 1;

        INSERT INTO dbo.PeriodoVacacionesHYT (TrabajadorId, FechaDesde, FechaHasta, DiasReconocidos)
        VALUES (@TrabajadorId, @FechaDesde, @FechaHasta, @DiasReconocidos);

        DECLARE @PeriodoId BIGINT = SCOPE_IDENTITY();
        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PeriodoId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PeriodoVacacionesHYT', @RegistroId, N'REGISTRAR';

        COMMIT TRANSACTION;
        SELECT @PeriodoId AS PeriodoVacacionesId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_VacacionesHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_VacacionesHYT_Guardar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_VacacionesHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @PeriodoVacacionesId BIGINT,
    @FechaInicio DATE,
    @FechaFin DATE,
    @Dias SMALLINT,
    @Estado VARCHAR(10),
    @VacacionesId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'VACACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar vacaciones.', 1;

        IF @FechaInicio IS NULL OR @FechaFin IS NULL OR @FechaFin < @FechaInicio
           OR @Dias IS NULL OR @Dias NOT BETWEEN 1 AND DATEDIFF(DAY, @FechaInicio, @FechaFin) + 1
           OR @Estado IS NULL OR @Estado NOT IN ('PENDIENTE', 'APROBADA', 'ANULADA')
           OR (@VacacionesId IS NULL AND @Estado = 'ANULADA')
            THROW 51000, N'Revise las fechas, los dias y el estado de las vacaciones.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND trabajador.Tipo = 'PLANILLA';

        IF @FechaIngreso IS NULL OR @FechaInicio < @FechaIngreso
            THROW 51000, N'Las vacaciones requieren un trabajador en planilla y fechas desde su ingreso.', 1;

        DECLARE @Reconocidos INT;
        SELECT @Reconocidos = periodo.DiasReconocidos
        FROM dbo.PeriodoVacacionesHYT AS periodo
        WHERE periodo.PeriodoVacacionesId = @PeriodoVacacionesId AND periodo.TrabajadorId = @TrabajadorId;

        IF @Reconocidos IS NULL
            THROW 51000, N'El periodo no corresponde al trabajador.', 1;

        DECLARE @EstadoAnterior VARCHAR(10);
        IF @VacacionesId IS NOT NULL
        BEGIN
            SELECT @EstadoAnterior = vacacion.Estado
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.VacacionesId = @VacacionesId AND vacacion.PeriodoVacacionesId = @PeriodoVacacionesId
              AND vacacion.VersionFila = @VersionFila AND vacacion.Estado <> 'ANULADA';

            IF @EstadoAnterior IS NULL
                THROW 51000, N'El registro cambio o esta anulado. Vuelva a abrirlo.', 1;
        END;

        IF @Estado <> 'ANULADA' AND EXISTS (
            SELECT 1
            FROM dbo.VacacionesHYT AS vacacion
            INNER JOIN dbo.PeriodoVacacionesHYT AS periodo ON periodo.PeriodoVacacionesId = vacacion.PeriodoVacacionesId
            WHERE periodo.TrabajadorId = @TrabajadorId AND vacacion.Estado <> 'ANULADA'
              AND (@VacacionesId IS NULL OR vacacion.VacacionesId <> @VacacionesId)
              AND vacacion.FechaInicio <= @FechaFin AND vacacion.FechaFin >= @FechaInicio)
            THROW 51000, N'Las fechas se superponen con otras vacaciones registradas.', 1;

        IF @Estado = 'APROBADA'
        BEGIN
            DECLARE @Asignados INT;
            SELECT @Asignados = COALESCE(SUM(CONVERT(INT, vacacion.Dias)), 0)
            FROM dbo.VacacionesHYT AS vacacion
            WHERE vacacion.PeriodoVacacionesId = @PeriodoVacacionesId AND vacacion.Estado = 'APROBADA'
              AND (@VacacionesId IS NULL OR vacacion.VacacionesId <> @VacacionesId);

            IF @Dias > @Reconocidos - @Asignados
                THROW 51000, N'Los dias exceden el saldo disponible del periodo.', 1;
        END;

        IF @VacacionesId IS NULL
        BEGIN
            INSERT INTO dbo.VacacionesHYT (PeriodoVacacionesId, FechaInicio, FechaFin, Dias, Estado)
            VALUES (@PeriodoVacacionesId, @FechaInicio, @FechaFin, @Dias, @Estado);
            SET @VacacionesId = SCOPE_IDENTITY();
        END
        ELSE
            UPDATE dbo.VacacionesHYT
            SET FechaInicio = CASE WHEN @Estado = 'ANULADA' THEN FechaInicio ELSE @FechaInicio END,
                FechaFin = CASE WHEN @Estado = 'ANULADA' THEN FechaFin ELSE @FechaFin END,
                Dias = CASE WHEN @Estado = 'ANULADA' THEN Dias ELSE @Dias END,
                Estado = @Estado
            WHERE VacacionesId = @VacacionesId;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @VacacionesId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'VacacionesHYT', @RegistroId,
             N'GUARDAR', @EstadoAnterior, @Estado;

        COMMIT TRANSACTION;
        SELECT @VacacionesId AS VacacionesId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_JornadaHYT_Listar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_JornadaHYT_Listar para el registro de asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_JornadaHYT_Listar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Desde DATE = NULL,
    @Hasta DATE = NULL,
    @JornadaId BIGINT = NULL,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'ASISTENCIA.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar asistencia.', 1;
        IF @Pagina IS NULL OR @Pagina < 1 OR @Desde > @Hasta
            THROW 51000, N'Pagina o rango de fechas no valido.', 1;

        SELECT jornada.JornadaId, jornada.Entrada, jornada.Salida, jornada.MinutosExtrasAprobados,
               jornada.VersionFila, COUNT_BIG(1) OVER () AS Total,
               SUM(CONVERT(BIGINT, jornada.MinutosExtrasAprobados)) OVER () AS MinutosExtrasTotal
        FROM dbo.JornadaHYT AS jornada
        WHERE jornada.TrabajadorId = @TrabajadorId
          AND (@JornadaId IS NULL OR jornada.JornadaId = @JornadaId)
          AND (@Desde IS NULL OR jornada.Entrada >= @Desde)
          AND (@Hasta IS NULL OR CONVERT(DATE, jornada.Entrada) <= @Hasta)
        ORDER BY jornada.Entrada DESC, jornada.JornadaId DESC
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: StoreProcedure/spu_JornadaHYT_Guardar.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_JornadaHYT_Guardar para el registro de asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_JornadaHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @TrabajadorId BIGINT,
    @Entrada DATETIME2(0),
    @Salida DATETIME2(0),
    @MinutosExtrasAprobados INT,
    @JornadaId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'ASISTENCIA.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar asistencia.', 1;
        IF @Entrada IS NULL OR @MinutosExtrasAprobados IS NULL OR @MinutosExtrasAprobados < 0
           OR (@Salida IS NULL AND @MinutosExtrasAprobados <> 0)
           OR (@Salida IS NOT NULL AND (@Salida <= @Entrada
               OR CONVERT(BIGINT, @MinutosExtrasAprobados) * 60 > DATEDIFF_BIG(SECOND, @Entrada, @Salida)))
            THROW 51000, N'Revise las horas de entrada, salida y minutos extras.', 1;

        BEGIN TRANSACTION;
        DECLARE @FechaIngreso DATE;
        SELECT @FechaIngreso = trabajador.FechaIngreso
        FROM dbo.TrabajadorHYT AS trabajador WITH (UPDLOCK, HOLDLOCK)
        WHERE trabajador.TrabajadorId = @TrabajadorId AND (trabajador.Activo = 1 OR @JornadaId IS NOT NULL);

        IF @FechaIngreso IS NULL OR @Entrada < @FechaIngreso
            THROW 51000, N'La jornada debe corresponder a un trabajador activo y a una fecha desde su ingreso.', 1;

        IF EXISTS (
            SELECT 1 FROM dbo.JornadaHYT AS jornada
            WHERE jornada.TrabajadorId = @TrabajadorId
              AND (@JornadaId IS NULL OR jornada.JornadaId <> @JornadaId)
              AND (@Salida IS NULL OR jornada.Entrada < @Salida)
              AND (jornada.Salida IS NULL OR jornada.Salida > @Entrada))
            THROW 51000, N'La jornada se superpone con otra jornada del trabajador.', 1;

        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @JornadaId IS NULL
        BEGIN
            INSERT INTO dbo.JornadaHYT (TrabajadorId, Entrada, Salida, MinutosExtrasAprobados)
            VALUES (@TrabajadorId, @Entrada, @Salida, @MinutosExtrasAprobados);
            SET @JornadaId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.JornadaHYT
            SET Entrada = @Entrada, Salida = @Salida, MinutosExtrasAprobados = @MinutosExtrasAprobados
            WHERE JornadaId = @JornadaId AND TrabajadorId = @TrabajadorId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La jornada cambio o no pertenece al trabajador. Vuelva a abrirla.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @JornadaId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Entrada AS Entrada, @Salida AS Salida, @MinutosExtrasAprobados AS MinutosExtrasAprobados
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'JornadaHYT', @RegistroId,
             @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @JornadaId AS JornadaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

-- Fuente: Script/033_PermisosPlanillas.sql
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra permisos y acceso a proveedores, costos, pagos, trabajadores, vacaciones y asistencia.
-- =============================================
USE [SistemaWebHYT];
GO
INSERT INTO dbo.Permiso (Codigo, Descripcion)
SELECT nuevo.Codigo, nuevo.Descripcion
FROM (VALUES
    (N'ASISTENCIA.CONSULTAR', N'Consultar jornadas y minutos extras aprobados'),
    (N'ASISTENCIA.EDITAR', N'Registrar jornadas y aprobar minutos extras'),
    (N'VACACIONES.CONSULTAR', N'Consultar los periodos y vacaciones de trabajadores'),
    (N'VACACIONES.EDITAR', N'Registrar periodos y gestionar vacaciones'),
    (N'TRABAJADORES.CONSULTAR', N'Consultar el maestro de trabajadores'),
    (N'TRABAJADORES.EDITAR', N'Crear y modificar trabajadores'),
    (N'TRABAJADORES.CUENTA', N'Consultar o cambiar la cuenta bancaria de trabajadores'),
    (N'PROVEEDORES.CONSULTAR', N'Consultar el maestro de proveedores'),
    (N'PROVEEDORES.EDITAR', N'Crear y modificar proveedores'),
    (N'PROVEEDORES.CUENTA', N'Consultar o cambiar la cuenta bancaria de proveedores'),
    (N'COSTOS_PROVEEDOR.CONSULTAR', N'Consultar costos de proveedores del presupuesto'),
    (N'COSTOS_PROVEEDOR.EDITAR', N'Registrar costos de proveedores del presupuesto'),
    (N'PAGOS_PROVEEDOR.CONSULTAR', N'Consultar saldos y pagos de proveedores'),
    (N'PAGOS_PROVEEDOR.EDITAR', N'Registrar pagos de proveedores')
) AS nuevo (Codigo, Descripcion)
WHERE NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WHERE permiso.Codigo = nuevo.Codigo);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_CostoProveedor_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PagoProveedor_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_CostoProveedor_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ProveedorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ProveedorHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PagoProveedor_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TrabajadorHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_TrabajadorHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_VacacionesHYT_Consultar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_PeriodoVacacionesHYT_Registrar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_VacacionesHYT_Guardar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_JornadaHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_JornadaHYT_Guardar TO hyt_app_executor;
GO

COMMIT TRANSACTION;
GO
