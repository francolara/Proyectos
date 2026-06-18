-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Despliega configuracion contable, tipos de comprobante SUNAT y modulo base de ventas.
-- =============================================

SET ANSI_NULLS ON;
SET QUOTED_IDENTIFIER ON;
GO

-- =============================================
-- Tablas de configuracion contable
-- =============================================

IF OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_ConfiguracionContabilizacion
    (
        IdConfiguracionContabilizacion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_ConfiguracionContabilizacion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        ModuloOperacion VARCHAR(10) NOT NULL,
        EscenarioOperacion VARCHAR(20) NOT NULL,
        IdOrigen INT NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        GeneraAsientoAutomatico BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_GeneraAsientoAutomatico DEFAULT (1),
        UsaTipoCambio BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_UsaTipoCambio DEFAULT (1),
        Activo BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacion_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_ModuloOperacion
            CHECK (ModuloOperacion IN ('COM', 'VEN'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacion_EscenarioOperacion
            CHECK (EscenarioOperacion IN ('MERCADERIA', 'GASTO', 'SERVICIO'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacion
        ADD CONSTRAINT UQ_CON_ConfiguracionContabilizacion
            UNIQUE (IdEmpresa, ModuloOperacion, EscenarioOperacion);
END;
GO

IF OBJECT_ID(N'dbo.CON_ConfiguracionContabilizacionDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_ConfiguracionContabilizacionDetalle
    (
        IdConfiguracionContabilizacionDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_ConfiguracionContabilizacionDetalle PRIMARY KEY,
        IdConfiguracionContabilizacion INT NOT NULL,
        Orden SMALLINT NOT NULL,
        ComponenteContable VARCHAR(20) NOT NULL,
        IdPlanCuenta INT NOT NULL,
        NaturalezaMovimiento CHAR(1) NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionDetalle_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_ConfiguracionContabilizacionDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionDetalle_Cabecera
            FOREIGN KEY (IdConfiguracionContabilizacion) REFERENCES dbo.CON_ConfiguracionContabilizacion (IdConfiguracionContabilizacion);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT FK_CON_ConfiguracionContabilizacionDetalle_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_Orden
            CHECK (Orden >= 1);

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_ComponenteContable
            CHECK (ComponenteContable IN ('BRUTO', 'IGV', 'TOTAL', 'REDONDEO', 'ISC', 'OTROS'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT CK_CON_ConfiguracionContabilizacionDetalle_NaturalezaMovimiento
            CHECK (NaturalezaMovimiento IN ('D', 'H'));

    ALTER TABLE dbo.CON_ConfiguracionContabilizacionDetalle
        ADD CONSTRAINT UQ_CON_ConfiguracionContabilizacionDetalle_Orden
            UNIQUE (IdConfiguracionContabilizacion, Orden);
END;
GO

-- =============================================
-- Maestro SUNAT de comprobantes
-- =============================================

IF OBJECT_ID(N'dbo.ADM_TipoComprobante', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_TipoComprobante
    (
        IdTipoComprobante INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_TipoComprobante PRIMARY KEY,
        CodigoTipoComprobante VARCHAR(3) NOT NULL,
        Descripcion NVARCHAR(150) NOT NULL,
        UsoCompras BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoCompras DEFAULT (0),
        UsoVentas BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_UsoVentas DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_ADM_TipoComprobante_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_TipoComprobante_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_TipoComprobante
        ADD CONSTRAINT UQ_ADM_TipoComprobante_Codigo UNIQUE (CodigoTipoComprobante);
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '01'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '01',
        N'Factura',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '03'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '03',
        N'Boleta de venta',
        0,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '07'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '07',
        N'Nota de credito',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '08'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '08',
        N'Nota de debito',
        1,
        1,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '14'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '14',
        N'Recibo por servicios publicos',
        1,
        0,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '50'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '50',
        N'Declaracion unica de aduanas',
        1,
        0,
        1,
        N'codex'
    );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM dbo.ADM_TipoComprobante AS t
    WHERE t.CodigoTipoComprobante = '91'
)
BEGIN
    INSERT INTO dbo.ADM_TipoComprobante
    (
        CodigoTipoComprobante,
        Descripcion,
        UsoCompras,
        UsoVentas,
        Estado,
        UsuarioRegistro
    )
    VALUES
    (
        '91',
        N'Comprobante de no domiciliado',
        1,
        0,
        1,
        N'codex'
    );
END;
GO

-- =============================================
-- Tablas de ventas
-- =============================================

IF OBJECT_ID(N'dbo.VEN_Venta', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VEN_Venta
    (
        IdVenta INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_VEN_Venta PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdCliente INT NOT NULL,
        IdConfiguracionContabilizacion INT NOT NULL,
        IdAsiento INT NULL,
        FechaEmision DATE NOT NULL,
        FechaContabilizacion DATE NOT NULL,
        TipoComprobante VARCHAR(3) NOT NULL,
        Serie VARCHAR(10) NOT NULL,
        Numero VARCHAR(20) NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18,6) NOT NULL CONSTRAINT DF_VEN_Venta_TipoCambio DEFAULT (1),
        BaseImponible DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_BaseImponible DEFAULT (0),
        Igv DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Igv DEFAULT (0),
        Isc DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Isc DEFAULT (0),
        OtrosTributos DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_OtrosTributos DEFAULT (0),
        Redondeo DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_Redondeo DEFAULT (0),
        ImporteTotal DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_Venta_ImporteTotal DEFAULT (0),
        Observacion NVARCHAR(500) NULL,
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_VEN_Venta_Estado DEFAULT (N'FACTURADO'),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_VEN_Venta_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_ADM_Cliente
            FOREIGN KEY (IdCliente) REFERENCES dbo.ADM_Cliente (IdCliente);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_CON_ConfiguracionContabilizacion
            FOREIGN KEY (IdConfiguracionContabilizacion) REFERENCES dbo.CON_ConfiguracionContabilizacion (IdConfiguracionContabilizacion);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT FK_VEN_Venta_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT CK_VEN_Venta_Montos
            CHECK (
                BaseImponible >= 0
                AND Igv >= 0
                AND Isc >= 0
                AND OtrosTributos >= 0
                AND Redondeo >= 0
                AND ImporteTotal >= 0
            );

    ALTER TABLE dbo.VEN_Venta
        ADD CONSTRAINT UQ_VEN_Venta_Documento
            UNIQUE (IdEmpresa, IdCliente, TipoComprobante, Serie, Numero);
END;
GO

IF OBJECT_ID(N'dbo.VEN_VentaDetalle', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.VEN_VentaDetalle
    (
        IdVentaDetalle INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_VEN_VentaDetalle PRIMARY KEY,
        IdVenta INT NOT NULL,
        Item SMALLINT NOT NULL,
        Descripcion NVARCHAR(250) NOT NULL,
        Cantidad DECIMAL(18,4) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_Cantidad DEFAULT (1),
        ValorUnitario DECIMAL(18,6) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_ValorUnitario DEFAULT (0),
        ImporteBruto DECIMAL(18,2) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_ImporteBruto DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_VEN_VentaDetalle_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT FK_VEN_VentaDetalle_VEN_Venta
            FOREIGN KEY (IdVenta) REFERENCES dbo.VEN_Venta (IdVenta);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT CK_VEN_VentaDetalle_Item
            CHECK (Item >= 1);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT CK_VEN_VentaDetalle_Montos
            CHECK (Cantidad > 0 AND ValorUnitario >= 0 AND ImporteBruto >= 0);

    ALTER TABLE dbo.VEN_VentaDetalle
        ADD CONSTRAINT UQ_VEN_VentaDetalle_Item
            UNIQUE (IdVenta, Item);
END;
GO

-- =============================================
-- Stored Procedures maestros
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarTiposComprobanteActivos
    @UsoCompras BIT = 0,
    @UsoVentas BIT = 0
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            t.IdTipoComprobante,
            t.CodigoTipoComprobante,
            t.Descripcion,
            t.UsoCompras,
            t.UsoVentas,
            t.Estado
        FROM dbo.ADM_TipoComprobante AS t
        WHERE t.Estado = 1
          AND (@UsoCompras = 0 OR t.UsoCompras = 1)
          AND (@UsoVentas = 0 OR t.UsoVentas = 1)
        ORDER BY
            t.CodigoTipoComprobante ASC;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_ADM_ListarClientesActivosPorEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdCliente,
            c.IdEmpresa,
            c.IdPersona,
            c.CodigoCliente,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto,
            c.LimiteCredito,
            c.DiasCredito,
            c.Observacion,
            c.Estado
        FROM dbo.ADM_Cliente AS c
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        WHERE c.IdEmpresa = @IdEmpresa
          AND c.Estado = 1
          AND pe.Estado = 1
        ORDER BY
            pe.NombreCompleto ASC;

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
GO

-- =============================================
-- Stored Procedures configuracion contable
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_ListarConfiguracionContabilizacionPorEmpresa
    @IdEmpresa INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdConfiguracionContabilizacion,
            c.IdEmpresa,
            c.ModuloOperacion,
            c.EscenarioOperacion,
            c.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            c.Descripcion,
            c.GeneraAsientoAutomatico,
            c.UsaTipoCambio,
            c.Activo,
            COUNT(d.IdConfiguracionContabilizacionDetalle) AS CantidadComponentes
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = c.IdOrigen
        LEFT JOIN dbo.CON_ConfiguracionContabilizacionDetalle AS d
            ON d.IdConfiguracionContabilizacion = c.IdConfiguracionContabilizacion
           AND d.Activo = 1
        WHERE c.IdEmpresa = @IdEmpresa
        GROUP BY
            c.IdConfiguracionContabilizacion,
            c.IdEmpresa,
            c.ModuloOperacion,
            c.EscenarioOperacion,
            c.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            c.Descripcion,
            c.GeneraAsientoAutomatico,
            c.UsaTipoCambio,
            c.Activo
        ORDER BY
            c.ModuloOperacion ASC,
            c.EscenarioOperacion ASC;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            c.IdConfiguracionContabilizacion,
            c.IdEmpresa,
            c.ModuloOperacion,
            c.EscenarioOperacion,
            c.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            c.Descripcion,
            c.GeneraAsientoAutomatico,
            c.UsaTipoCambio,
            c.Activo
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        INNER JOIN dbo.CON_Origen AS o
            ON o.IdOrigen = c.IdOrigen
        WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        SELECT
            d.IdConfiguracionContabilizacionDetalle,
            d.IdConfiguracionContabilizacion,
            d.Orden,
            d.ComponenteContable,
            d.IdPlanCuenta,
            p.CodigoCuenta,
            p.NombreCuenta,
            d.NaturalezaMovimiento,
            d.Activo
        FROM dbo.CON_ConfiguracionContabilizacionDetalle AS d
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = d.IdPlanCuenta
        WHERE d.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
        ORDER BY
            d.Orden ASC;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT = NULL,
    @IdEmpresa INT,
    @ModuloOperacion VARCHAR(10),
    @EscenarioOperacion VARCHAR(20),
    @IdOrigen INT,
    @Descripcion NVARCHAR(200),
    @GeneraAsientoAutomatico BIT,
    @UsaTipoCambio BIT,
    @Activo BIT,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdConfiguracionTrabajo INT

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la configuracion contable.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdOrigen = @IdOrigen
              AND o.IdEmpresa = @IdEmpresa
              AND o.Estado = 1
        )
        BEGIN
            RAISERROR(N'El origen indicado no existe o no pertenece a la empresa.', 16, 1);
        END;

        DECLARE @Detalle TABLE
        (
            Orden SMALLINT NOT NULL,
            ComponenteContable VARCHAR(20) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            NaturalezaMovimiento CHAR(1) NOT NULL,
            Activo BIT NOT NULL
        );

        INSERT INTO @Detalle
        (
            Orden,
            ComponenteContable,
            IdPlanCuenta,
            NaturalezaMovimiento,
            Activo
        )
        SELECT
            T.N.value('@Orden', 'smallint'),
            T.N.value('@ComponenteContable', 'varchar(20)'),
            T.N.value('@IdPlanCuenta', 'int'),
            T.N.value('@NaturalezaMovimiento', 'char(1)'),
            T.N.value('@Activo', 'bit')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos un componente contable.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.ComponenteContable
            FROM @Detalle AS d
            WHERE d.Activo = 1
            GROUP BY
                d.ComponenteContable
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten componentes activos duplicados en la misma configuracion.', 16, 1);
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
            RAISERROR(N'Todas las cuentas configuradas deben existir, estar activas y aceptar movimiento.', 16, 1);
        END;

        BEGIN TRAN;

        IF @IdConfiguracionContabilizacion IS NULL
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_ConfiguracionContabilizacion AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.ModuloOperacion = @ModuloOperacion
                  AND c.EscenarioOperacion = @EscenarioOperacion
            )
            BEGIN
                RAISERROR(N'Ya existe una configuracion para la empresa, modulo y escenario seleccionados.', 16, 1);
            END;

            INSERT INTO dbo.CON_ConfiguracionContabilizacion
            (
                IdEmpresa,
                ModuloOperacion,
                EscenarioOperacion,
                IdOrigen,
                Descripcion,
                GeneraAsientoAutomatico,
                UsaTipoCambio,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @ModuloOperacion,
                @EscenarioOperacion,
                @IdOrigen,
                @Descripcion,
                @GeneraAsientoAutomatico,
                @UsaTipoCambio,
                @Activo,
                @UsuarioRegistro
            );

            SET @IdConfiguracionTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SET @IdConfiguracionTrabajo = @IdConfiguracionContabilizacion;

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_ConfiguracionContabilizacion AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.ModuloOperacion = @ModuloOperacion
                  AND c.EscenarioOperacion = @EscenarioOperacion
                  AND c.IdConfiguracionContabilizacion <> @IdConfiguracionContabilizacion
            )
            BEGIN
                RAISERROR(N'Ya existe otra configuracion para la empresa, modulo y escenario seleccionados.', 16, 1);
            END;

            UPDATE dbo.CON_ConfiguracionContabilizacion
            SET ModuloOperacion = @ModuloOperacion,
                EscenarioOperacion = @EscenarioOperacion,
                IdOrigen = @IdOrigen,
                Descripcion = @Descripcion,
                GeneraAsientoAutomatico = @GeneraAsientoAutomatico,
                UsaTipoCambio = @UsaTipoCambio,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
              AND IdEmpresa = @IdEmpresa;

            IF @@ROWCOUNT = 0
            BEGIN
                RAISERROR(N'La configuracion indicada no existe para la empresa activa.', 16, 1);
            END;

            DELETE FROM dbo.CON_ConfiguracionContabilizacionDetalle
            WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;
        END;

        INSERT INTO dbo.CON_ConfiguracionContabilizacionDetalle
        (
            IdConfiguracionContabilizacion,
            Orden,
            ComponenteContable,
            IdPlanCuenta,
            NaturalezaMovimiento,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdConfiguracionTrabajo,
            d.Orden,
            d.ComponenteContable,
            d.IdPlanCuenta,
            d.NaturalezaMovimiento,
            d.Activo,
            @UsuarioRegistro
        FROM @Detalle AS d
        ORDER BY
            d.Orden ASC;

        COMMIT;

        SELECT
            @IdConfiguracionTrabajo AS IdConfiguracionContabilizacion;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        BEGIN TRAN;

        DELETE FROM dbo.CON_ConfiguracionContabilizacionDetalle
        WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        DELETE FROM dbo.CON_ConfiguracionContabilizacion
        WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;

        COMMIT;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

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
GO

-- =============================================
-- Stored Procedures ventas
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ListarVentasPorEmpresa
    @IdEmpresa INT,
    @Periodo CHAR(6) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            v.IdVenta,
            v.IdEmpresa,
            v.IdCliente,
            c.CodigoCliente,
            pe.NombreCompleto AS NombreCliente,
            v.IdConfiguracionContabilizacion,
            cfg.ModuloOperacion,
            cfg.EscenarioOperacion,
            v.IdAsiento,
            v.FechaEmision,
            v.FechaContabilizacion,
            CONVERT(CHAR(6), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) AS Periodo,
            v.TipoComprobante,
            v.Serie,
            v.Numero,
            v.IdMoneda,
            m.CodigoMoneda,
            v.TipoCambio,
            v.BaseImponible,
            v.Igv,
            v.Isc,
            v.OtrosTributos,
            v.Redondeo,
            v.ImporteTotal,
            v.Observacion,
            v.Estado
        FROM dbo.VEN_Venta AS v
        INNER JOIN dbo.ADM_Cliente AS c
            ON c.IdCliente = v.IdCliente
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
            ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = v.IdMoneda
        WHERE v.IdEmpresa = @IdEmpresa
          AND (
                @Periodo IS NULL
                OR CONVERT(CHAR(6), YEAR(v.FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(v.FechaContabilizacion)), 2) = @Periodo
              )
        ORDER BY
            v.FechaContabilizacion DESC,
            v.IdVenta DESC;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_VEN_ObtenerVenta
    @IdVenta INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT
            v.IdVenta,
            v.IdEmpresa,
            v.IdCliente,
            c.CodigoCliente,
            pe.TipoDocumento,
            pe.NumeroDocumento,
            pe.NombreCompleto AS NombreCliente,
            v.IdConfiguracionContabilizacion,
            cfg.ModuloOperacion,
            cfg.EscenarioOperacion,
            cfg.Descripcion AS DescripcionConfiguracion,
            v.IdAsiento,
            v.FechaEmision,
            v.FechaContabilizacion,
            v.TipoComprobante,
            v.Serie,
            v.Numero,
            v.IdMoneda,
            m.CodigoMoneda,
            v.TipoCambio,
            v.BaseImponible,
            v.Igv,
            v.Isc,
            v.OtrosTributos,
            v.Redondeo,
            v.ImporteTotal,
            v.Observacion,
            v.Estado
        FROM dbo.VEN_Venta AS v
        INNER JOIN dbo.ADM_Cliente AS c
            ON c.IdCliente = v.IdCliente
        INNER JOIN dbo.ADM_Persona AS pe
            ON pe.IdPersona = c.IdPersona
        INNER JOIN dbo.CON_ConfiguracionContabilizacion AS cfg
            ON cfg.IdConfiguracionContabilizacion = v.IdConfiguracionContabilizacion
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = v.IdMoneda
        WHERE v.IdVenta = @IdVenta;

        SELECT
            d.IdVentaDetalle,
            d.IdVenta,
            d.Item,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto
        FROM dbo.VEN_VentaDetalle AS d
        WHERE d.IdVenta = @IdVenta
        ORDER BY
            d.Item ASC;

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
GO

CREATE OR ALTER PROCEDURE dbo.usp_VEN_GuardarVentaConAsiento
    @IdVenta INT = NULL,
    @IdEmpresa INT,
    @IdCliente INT,
    @IdConfiguracionContabilizacion INT,
    @FechaEmision DATE,
    @FechaContabilizacion DATE,
    @TipoComprobante VARCHAR(3),
    @Serie VARCHAR(10),
    @Numero VARCHAR(20),
    @IdMoneda INT,
    @TipoCambio DECIMAL(18,6),
    @BaseImponible DECIMAL(18,2),
    @Igv DECIMAL(18,2),
    @Isc DECIMAL(18,2),
    @OtrosTributos DECIMAL(18,2),
    @Redondeo DECIMAL(18,2),
    @ImporteTotal DECIMAL(18,2),
    @Observacion NVARCHAR(500) = NULL,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdVentaTrabajo INT
        DECLARE @IdAsientoTrabajo INT
        DECLARE @IdOrigen INT
        DECLARE @Periodo CHAR(6) = CONVERT(CHAR(4), YEAR(@FechaContabilizacion)) + RIGHT('0' + CONVERT(VARCHAR(2), MONTH(@FechaContabilizacion)), 2)
        DECLARE @Ejercicio SMALLINT = YEAR(@FechaContabilizacion)
        DECLARE @Mes TINYINT = MONTH(@FechaContabilizacion)
        DECLARE @NumeroAsiento INT
        DECLARE @GlosaAsiento NVARCHAR(500)
        DECLARE @TotalDebe DECIMAL(18,2)
        DECLARE @TotalHaber DECIMAL(18,2)
        DECLARE @EstadoConfiguracion BIT
        DECLARE @GeneraAsientoAutomatico BIT

        IF @BaseImponible < 0
           OR @Igv < 0
           OR @Isc < 0
           OR @OtrosTributos < 0
           OR @Redondeo < 0
           OR @ImporteTotal < 0
        BEGIN
            RAISERROR(N'Los montos de la venta no pueden ser negativos.', 16, 1);
        END;

        IF @ImporteTotal <> (@BaseImponible + @Igv + @Isc + @OtrosTributos + @Redondeo)
        BEGIN
            RAISERROR(N'El importe total debe ser igual a la suma de bruto, IGV, ISC, otros tributos y redondeo.', 16, 1);
        END;

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la venta.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Cliente AS c
            WHERE c.IdCliente = @IdCliente
              AND c.IdEmpresa = @IdEmpresa
              AND c.Estado = 1
        )
        BEGIN
            RAISERROR(N'El cliente seleccionado no existe o no pertenece a la empresa.', 16, 1);
        END;

        SELECT
            @IdOrigen = c.IdOrigen,
            @EstadoConfiguracion = c.Activo,
            @GeneraAsientoAutomatico = c.GeneraAsientoAutomatico
        FROM dbo.CON_ConfiguracionContabilizacion AS c
        WHERE c.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
          AND c.IdEmpresa = @IdEmpresa
          AND c.ModuloOperacion = 'VEN';

        IF @IdOrigen IS NULL
        BEGIN
            RAISERROR(N'La configuracion contable indicada no existe para ventas en la empresa activa.', 16, 1);
        END;

        IF @EstadoConfiguracion = 0
        BEGIN
            RAISERROR(N'La configuracion contable seleccionada esta inactiva.', 16, 1);
        END;

        IF @GeneraAsientoAutomatico = 0
        BEGIN
            RAISERROR(N'La configuracion seleccionada no esta habilitada para generar asiento automatico.', 16, 1);
        END;

        DECLARE @DetalleVenta TABLE
        (
            Item SMALLINT NOT NULL,
            Descripcion NVARCHAR(250) NOT NULL,
            Cantidad DECIMAL(18,4) NOT NULL,
            ValorUnitario DECIMAL(18,6) NOT NULL,
            ImporteBruto DECIMAL(18,2) NOT NULL
        );

        INSERT INTO @DetalleVenta
        (
            Item,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto
        )
        SELECT
            T.N.value('@Item', 'smallint'),
            T.N.value('@Descripcion', 'nvarchar(250)'),
            T.N.value('@Cantidad', 'decimal(18,4)'),
            T.N.value('@ValorUnitario', 'decimal(18,6)'),
            T.N.value('@ImporteBruto', 'decimal(18,2)')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @DetalleVenta
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos una linea en la venta.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @DetalleVenta AS d
            WHERE d.Item < 1
               OR d.Cantidad <= 0
               OR d.ValorUnitario < 0
               OR d.ImporteBruto < 0
        )
        BEGIN
            RAISERROR(N'El detalle de la venta contiene valores no validos.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.Item
            FROM @DetalleVenta AS d
            GROUP BY
                d.Item
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten items duplicados en el detalle de la venta.', 16, 1);
        END;

        DECLARE @AsientoDetalle TABLE
        (
            Item SMALLINT IDENTITY(1,1) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            Debe DECIMAL(18,2) NOT NULL,
            Haber DECIMAL(18,2) NOT NULL,
            GlosaDetalle NVARCHAR(300) NULL
        );

        INSERT INTO @AsientoDetalle
        (
            IdPlanCuenta,
            Debe,
            Haber,
            GlosaDetalle
        )
        SELECT
            d.IdPlanCuenta,
            CASE d.NaturalezaMovimiento
                WHEN 'D' THEN
                    CASE d.ComponenteContable
                        WHEN 'BRUTO' THEN @BaseImponible
                        WHEN 'IGV' THEN @Igv
                        WHEN 'ISC' THEN @Isc
                        WHEN 'OTROS' THEN @OtrosTributos
                        WHEN 'REDONDEO' THEN @Redondeo
                        WHEN 'TOTAL' THEN @ImporteTotal
                        ELSE 0
                    END
                ELSE 0
            END AS Debe,
            CASE d.NaturalezaMovimiento
                WHEN 'H' THEN
                    CASE d.ComponenteContable
                        WHEN 'BRUTO' THEN @BaseImponible
                        WHEN 'IGV' THEN @Igv
                        WHEN 'ISC' THEN @Isc
                        WHEN 'OTROS' THEN @OtrosTributos
                        WHEN 'REDONDEO' THEN @Redondeo
                        WHEN 'TOTAL' THEN @ImporteTotal
                        ELSE 0
                    END
                ELSE 0
            END AS Haber,
            CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero, N' / ', d.ComponenteContable) AS GlosaDetalle
        FROM dbo.CON_ConfiguracionContabilizacionDetalle AS d
        WHERE d.IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
          AND d.Activo = 1;

        DELETE FROM @AsientoDetalle
        WHERE Debe = 0
          AND Haber = 0;

        IF NOT EXISTS
        (
            SELECT 1
            FROM @AsientoDetalle
        )
        BEGIN
            RAISERROR(N'La configuracion seleccionada no genera lineas contables con los importes de la venta.', 16, 1);
        END;

        SELECT
            @TotalDebe = SUM(d.Debe),
            @TotalHaber = SUM(d.Haber)
        FROM @AsientoDetalle AS d;

        IF @TotalDebe <> @TotalHaber
        BEGIN
            RAISERROR(N'La configuracion contable de ventas no genera un asiento cuadrado para los importes ingresados.', 16, 1);
        END;

        SET @GlosaAsiento = CONCAT(N'Venta ', @TipoComprobante, N' ', @Serie, N'-', @Numero);

        SET TRANSACTION ISOLATION LEVEL SERIALIZABLE;

        BEGIN TRAN;

        IF @IdVenta IS NULL
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
                @FechaContabilizacion,
                @GlosaAsiento,
                @IdMoneda,
                @TipoCambio,
                @TotalDebe,
                @TotalHaber,
                N'FACTURADO',
                CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                @Observacion,
                @UsuarioRegistro
            );

            SET @IdAsientoTrabajo = SCOPE_IDENTITY();

            INSERT INTO dbo.VEN_Venta
            (
                IdEmpresa,
                IdCliente,
                IdConfiguracionContabilizacion,
                IdAsiento,
                FechaEmision,
                FechaContabilizacion,
                TipoComprobante,
                Serie,
                Numero,
                IdMoneda,
                TipoCambio,
                BaseImponible,
                Igv,
                Isc,
                OtrosTributos,
                Redondeo,
                ImporteTotal,
                Observacion,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdCliente,
                @IdConfiguracionContabilizacion,
                @IdAsientoTrabajo,
                @FechaEmision,
                @FechaContabilizacion,
                @TipoComprobante,
                @Serie,
                @Numero,
                @IdMoneda,
                @TipoCambio,
                @BaseImponible,
                @Igv,
                @Isc,
                @OtrosTributos,
                @Redondeo,
                @ImporteTotal,
                @Observacion,
                N'FACTURADO',
                @UsuarioRegistro
            );

            SET @IdVentaTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SELECT
                @IdVentaTrabajo = v.IdVenta,
                @IdAsientoTrabajo = v.IdAsiento
            FROM dbo.VEN_Venta AS v
            WHERE v.IdVenta = @IdVenta
              AND v.IdEmpresa = @IdEmpresa;

            IF @IdVentaTrabajo IS NULL
            BEGIN
                RAISERROR(N'La venta indicada no existe para la empresa activa.', 16, 1);
            END;

            UPDATE dbo.VEN_Venta
            SET IdCliente = @IdCliente,
                IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion,
                FechaEmision = @FechaEmision,
                FechaContabilizacion = @FechaContabilizacion,
                TipoComprobante = @TipoComprobante,
                Serie = @Serie,
                Numero = @Numero,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                BaseImponible = @BaseImponible,
                Igv = @Igv,
                Isc = @Isc,
                OtrosTributos = @OtrosTributos,
                Redondeo = @Redondeo,
                ImporteTotal = @ImporteTotal,
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdVenta = @IdVentaTrabajo;

            DELETE FROM dbo.CON_AsientoDetalle
            WHERE IdAsiento = @IdAsientoTrabajo;

            UPDATE dbo.CON_Asiento
            SET FechaAsiento = @FechaContabilizacion,
                Glosa = @GlosaAsiento,
                IdMoneda = @IdMoneda,
                TipoCambio = @TipoCambio,
                TotalDebe = @TotalDebe,
                TotalHaber = @TotalHaber,
                Estado = N'FACTURADO',
                ReferenciaExterna = CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
                Observacion = @Observacion,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdAsiento = @IdAsientoTrabajo;

            DELETE FROM dbo.VEN_VentaDetalle
            WHERE IdVenta = @IdVentaTrabajo;
        END;

        INSERT INTO dbo.CON_AsientoDetalle
        (
            IdAsiento,
            Item,
            IdPlanCuenta,
            GlosaDetalle,
            IdCliente,
            Debe,
            Haber,
            ReferenciaLinea,
            UsuarioRegistro
        )
        SELECT
            @IdAsientoTrabajo,
            d.Item,
            d.IdPlanCuenta,
            d.GlosaDetalle,
            @IdCliente,
            d.Debe,
            d.Haber,
            CONCAT(@TipoComprobante, N' ', @Serie, N'-', @Numero),
            @UsuarioRegistro
        FROM @AsientoDetalle AS d
        ORDER BY
            d.Item ASC;

        INSERT INTO dbo.VEN_VentaDetalle
        (
            IdVenta,
            Item,
            Descripcion,
            Cantidad,
            ValorUnitario,
            ImporteBruto,
            UsuarioRegistro
        )
        SELECT
            @IdVentaTrabajo,
            d.Item,
            d.Descripcion,
            d.Cantidad,
            d.ValorUnitario,
            d.ImporteBruto,
            @UsuarioRegistro
        FROM @DetalleVenta AS d
        ORDER BY
            d.Item ASC;

        COMMIT;

        SET TRANSACTION ISOLATION LEVEL READ COMMITTED;

        SELECT
            v.IdVenta,
            v.IdAsiento,
            v.ImporteTotal,
            v.Estado
        FROM dbo.VEN_Venta AS v
        WHERE v.IdVenta = @IdVentaTrabajo;

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
GO
