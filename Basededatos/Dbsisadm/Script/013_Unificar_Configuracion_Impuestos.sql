-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Unifica configuracion de impuestos en CON_TipoImpuestoConfiguracionEmpresa.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | Conserva CodigoCuenta en el maestro de impuestos y IdPlanCuenta solo en la configuracion por empresa.

IF COL_LENGTH(N'dbo.CON_TipoImpuesto', N'CodigoCuenta') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuesto
        ADD CodigoCuenta VARCHAR(20) NULL;
END;

IF OBJECT_ID(N'dbo.CON_TipoImpuestoConfiguracionEmpresa', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
    (
        IdTipoImpuestoConfiguracionEmpresa INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoImpuestoConfiguracionEmpresa PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdTipoImpuesto INT NOT NULL,
        IdPlanCuenta INT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_TipoImpuestoConfiguracionEmpresa_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_TipoImpuestoConfiguracionEmpresa_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );
END;

IF OBJECT_ID(N'dbo.FK_CON_TipoImpuestoConfiguracionEmpresa_SEG_Empresa', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);
END;

IF OBJECT_ID(N'dbo.FK_CON_TipoImpuestoConfiguracionEmpresa_CON_TipoImpuesto', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_TipoImpuesto
            FOREIGN KEY (IdTipoImpuesto) REFERENCES dbo.CON_TipoImpuesto (IdTipoImpuesto);
END;

IF OBJECT_ID(N'dbo.FK_CON_TipoImpuestoConfiguracionEmpresa_CON_PlanCuenta', N'F') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT FK_CON_TipoImpuestoConfiguracionEmpresa_CON_PlanCuenta
            FOREIGN KEY (IdPlanCuenta) REFERENCES dbo.CON_PlanCuenta (IdPlanCuenta);
END;

IF OBJECT_ID(N'dbo.UQ_CON_TipoImpuestoConfiguracionEmpresa', N'UQ') IS NULL
BEGIN
    ALTER TABLE dbo.CON_TipoImpuestoConfiguracionEmpresa
        ADD CONSTRAINT UQ_CON_TipoImpuestoConfiguracionEmpresa
            UNIQUE (IdEmpresa, IdTipoImpuesto);
END;

IF OBJECT_ID(N'dbo.CON_TipoImpuesto_Compra', N'U') IS NOT NULL
   OR OBJECT_ID(N'dbo.CON_TipoImpuesto_Venta', N'U') IS NOT NULL
BEGIN
    CREATE TABLE #ConfiguracionImpuestoOrigen
    (
        IdEmpresa INT NOT NULL,
        IdTipoImpuesto INT NOT NULL,
        IdPlanCuenta INT NULL,
        Activo BIT NOT NULL,
        UsuarioRegistro NVARCHAR(450) NULL,
        Prioridad INT NOT NULL
    );

    IF OBJECT_ID(N'dbo.CON_TipoImpuesto_Compra', N'U') IS NOT NULL
    BEGIN
        INSERT INTO #ConfiguracionImpuestoOrigen
        (
            IdEmpresa,
            IdTipoImpuesto,
            IdPlanCuenta,
            Activo,
            UsuarioRegistro,
            Prioridad
        )
        SELECT
            c.IdEmpresa,
            c.IdTipoImpuesto,
            COALESCE(c.IdPlanCuentaSoles, c.IdPlanCuentaDolares) AS IdPlanCuenta,
            c.Activo,
            c.UsuarioRegistro,
            1 AS Prioridad
        FROM dbo.CON_TipoImpuesto_Compra AS c;
    END;

    IF OBJECT_ID(N'dbo.CON_TipoImpuesto_Venta', N'U') IS NOT NULL
    BEGIN
        INSERT INTO #ConfiguracionImpuestoOrigen
        (
            IdEmpresa,
            IdTipoImpuesto,
            IdPlanCuenta,
            Activo,
            UsuarioRegistro,
            Prioridad
        )
        SELECT
            v.IdEmpresa,
            v.IdTipoImpuesto,
            COALESCE(v.IdPlanCuentaSoles, v.IdPlanCuentaDolares) AS IdPlanCuenta,
            v.Activo,
            v.UsuarioRegistro,
            2 AS Prioridad
        FROM dbo.CON_TipoImpuesto_Venta AS v;
    END;

    ;WITH
    ConfiguracionElegida AS
    (
        SELECT
            o.IdEmpresa,
            o.IdTipoImpuesto,
            o.IdPlanCuenta,
            o.Activo,
            o.UsuarioRegistro,
            ROW_NUMBER() OVER (PARTITION BY o.IdEmpresa, o.IdTipoImpuesto ORDER BY o.Prioridad ASC) AS Fila
        FROM #ConfiguracionImpuestoOrigen AS o
    )
    MERGE dbo.CON_TipoImpuestoConfiguracionEmpresa AS destino
    USING
    (
        SELECT
            e.IdEmpresa,
            e.IdTipoImpuesto,
            e.IdPlanCuenta,
            e.Activo,
            e.UsuarioRegistro
        FROM ConfiguracionElegida AS e
        WHERE e.Fila = 1
    ) AS fuente
    ON destino.IdEmpresa = fuente.IdEmpresa
       AND destino.IdTipoImpuesto = fuente.IdTipoImpuesto
    WHEN MATCHED THEN
        UPDATE SET
            IdPlanCuenta = COALESCE(destino.IdPlanCuenta, fuente.IdPlanCuenta),
            Activo = fuente.Activo,
            UsuarioRegistro = fuente.UsuarioRegistro
    WHEN NOT MATCHED BY TARGET THEN
        INSERT
        (
            IdEmpresa,
            IdTipoImpuesto,
            IdPlanCuenta,
            Activo,
            UsuarioRegistro
        )
        VALUES
        (
            fuente.IdEmpresa,
            fuente.IdTipoImpuesto,
            fuente.IdPlanCuenta,
            fuente.Activo,
            fuente.UsuarioRegistro
        );

    DROP TABLE #ConfiguracionImpuestoOrigen;
END;

IF OBJECT_ID(N'dbo.CON_TipoImpuesto_Compra', N'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.CON_TipoImpuesto_Compra;
END;

IF OBJECT_ID(N'dbo.CON_TipoImpuesto_Venta', N'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.CON_TipoImpuesto_Venta;
END;
