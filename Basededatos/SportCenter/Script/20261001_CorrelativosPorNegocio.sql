
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/10/2026
-- Description:   Agrega la estructura para correlativos visibles e independientes por negocio.
--                Completa correlativos historicos y no utiliza triggers.
-- =============================================
-- Firma:         FRANCO LARA - 08/10/2026 | Completa correlativos historicos por negocio y sincroniza el ultimo numero para evitar valores nulos en los listados.
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    IF OBJECT_ID(N'dbo.NegocioCorrelativos', N'U') IS NULL
    BEGIN
        CREATE TABLE dbo.NegocioCorrelativos
        (
            NegocioId INT NOT NULL,
            Entidad NVARCHAR(30) NOT NULL,
            UltimoNumero INT NOT NULL,
            FechaActualizacion DATETIME2(7) NOT NULL
                CONSTRAINT DF_NegocioCorrelativos_FechaActualizacion DEFAULT (SYSUTCDATETIME()),
            UsuarioActualizacion NVARCHAR(200) NULL,
            CONSTRAINT PK_NegocioCorrelativos
                PRIMARY KEY CLUSTERED (NegocioId ASC, Entidad ASC),
            CONSTRAINT CK_NegocioCorrelativos_Entidad
                CHECK (Entidad IN (N'RESERVA', N'PAGO', N'CLIENTE', N'SEDE', N'ESPACIO', N'CUPON', N'PROMOCION')),
            CONSTRAINT CK_NegocioCorrelativos_UltimoNumero
                CHECK (UltimoNumero >= 0),
            CONSTRAINT FK_NegocioCorrelativos_Negocios_NegocioId
                FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id) ON DELETE CASCADE
        );
    END;

    IF COL_LENGTH(N'dbo.Reservas', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Reservas ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Pagos', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Pagos ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Clientes', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Clientes ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Sedes', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Sedes ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.EspaciosDeportivos ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Cupones', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Cupones ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.PromocionesHorario', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.PromocionesHorario ADD NumeroPorNegocio INT NULL;

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Sedes s
        WHERE s.NumeroPorNegocio > 0
        GROUP BY s.NegocioId, s.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en Sedes.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.NumeroPorNegocio > 0
        GROUP BY s.NegocioId, e.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en EspaciosDeportivos.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Clientes c
        WHERE c.NumeroPorNegocio > 0
        GROUP BY c.NegocioId, c.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en Clientes.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE r.NumeroPorNegocio > 0
        GROUP BY s.NegocioId, r.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en Reservas.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE p.NumeroPorNegocio > 0
        GROUP BY s.NegocioId, p.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en Pagos.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.Cupones c
        WHERE c.NumeroPorNegocio > 0
        GROUP BY c.NegocioId, c.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en Cupones.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM dbo.PromocionesHorario p
        WHERE p.NumeroPorNegocio > 0
        GROUP BY p.NegocioId, p.NumeroPorNegocio
        HAVING COUNT_BIG(1) > 1
    )
        RAISERROR('Existen correlativos duplicados en PromocionesHorario.', 16, 1);

    ;WITH Maximos AS
    (
        SELECT s.NegocioId, MAX(s.NumeroPorNegocio) AS UltimoNumero
        FROM dbo.Sedes s
        WHERE s.NumeroPorNegocio > 0
        GROUP BY s.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            s.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY s.NegocioId ORDER BY s.Id) AS NumeroPorNegocio
        FROM dbo.Sedes s
        LEFT JOIN Maximos m ON m.NegocioId = s.NegocioId
        WHERE s.NumeroPorNegocio IS NULL OR s.NumeroPorNegocio <= 0
    )
    UPDATE s
    SET NumeroPorNegocio = CONVERT(INT, p.NumeroPorNegocio)
    FROM dbo.Sedes s
    INNER JOIN Pendientes p ON p.Id = s.Id;

    ;WITH Base AS
    (
        SELECT e.Id, s.NegocioId, e.NumeroPorNegocio
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    ),
    Maximos AS
    (
        SELECT b.NegocioId, MAX(b.NumeroPorNegocio) AS UltimoNumero
        FROM Base b
        WHERE b.NumeroPorNegocio > 0
        GROUP BY b.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            b.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY b.NegocioId ORDER BY b.Id) AS NumeroPorNegocio
        FROM Base b
        LEFT JOIN Maximos m ON m.NegocioId = b.NegocioId
        WHERE b.NumeroPorNegocio IS NULL OR b.NumeroPorNegocio <= 0
    )
    UPDATE e
    SET NumeroPorNegocio = CONVERT(INT, p.NumeroPorNegocio)
    FROM dbo.EspaciosDeportivos e
    INNER JOIN Pendientes p ON p.Id = e.Id;

    ;WITH Maximos AS
    (
        SELECT c.NegocioId, MAX(c.NumeroPorNegocio) AS UltimoNumero
        FROM dbo.Clientes c
        WHERE c.NumeroPorNegocio > 0
        GROUP BY c.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            c.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY c.NegocioId ORDER BY c.Id) AS NumeroPorNegocio
        FROM dbo.Clientes c
        LEFT JOIN Maximos m ON m.NegocioId = c.NegocioId
        WHERE c.NumeroPorNegocio IS NULL OR c.NumeroPorNegocio <= 0
    )
    UPDATE c
    SET NumeroPorNegocio = CONVERT(INT, p.NumeroPorNegocio)
    FROM dbo.Clientes c
    INNER JOIN Pendientes p ON p.Id = c.Id;

    ;WITH Base AS
    (
        SELECT r.Id, s.NegocioId, r.NumeroPorNegocio
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    ),
    Maximos AS
    (
        SELECT b.NegocioId, MAX(b.NumeroPorNegocio) AS UltimoNumero
        FROM Base b
        WHERE b.NumeroPorNegocio > 0
        GROUP BY b.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            b.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY b.NegocioId ORDER BY b.Id) AS NumeroPorNegocio
        FROM Base b
        LEFT JOIN Maximos m ON m.NegocioId = b.NegocioId
        WHERE b.NumeroPorNegocio IS NULL OR b.NumeroPorNegocio <= 0
    )
    UPDATE r
    SET NumeroPorNegocio = CONVERT(INT, p.NumeroPorNegocio)
    FROM dbo.Reservas r
    INNER JOIN Pendientes p ON p.Id = r.Id;

    ;WITH Base AS
    (
        SELECT p.Id, s.NegocioId, p.NumeroPorNegocio
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    ),
    Maximos AS
    (
        SELECT b.NegocioId, MAX(b.NumeroPorNegocio) AS UltimoNumero
        FROM Base b
        WHERE b.NumeroPorNegocio > 0
        GROUP BY b.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            b.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY b.NegocioId ORDER BY b.Id) AS NumeroPorNegocio
        FROM Base b
        LEFT JOIN Maximos m ON m.NegocioId = b.NegocioId
        WHERE b.NumeroPorNegocio IS NULL OR b.NumeroPorNegocio <= 0
    )
    UPDATE p
    SET NumeroPorNegocio = CONVERT(INT, x.NumeroPorNegocio)
    FROM dbo.Pagos p
    INNER JOIN Pendientes x ON x.Id = p.Id;

    ;WITH Maximos AS
    (
        SELECT c.NegocioId, MAX(c.NumeroPorNegocio) AS UltimoNumero
        FROM dbo.Cupones c
        WHERE c.NumeroPorNegocio > 0
        GROUP BY c.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            c.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY c.NegocioId ORDER BY c.Id) AS NumeroPorNegocio
        FROM dbo.Cupones c
        LEFT JOIN Maximos m ON m.NegocioId = c.NegocioId
        WHERE c.NumeroPorNegocio IS NULL OR c.NumeroPorNegocio <= 0
    )
    UPDATE c
    SET NumeroPorNegocio = CONVERT(INT, p.NumeroPorNegocio)
    FROM dbo.Cupones c
    INNER JOIN Pendientes p ON p.Id = c.Id;

    ;WITH Maximos AS
    (
        SELECT p.NegocioId, MAX(p.NumeroPorNegocio) AS UltimoNumero
        FROM dbo.PromocionesHorario p
        WHERE p.NumeroPorNegocio > 0
        GROUP BY p.NegocioId
    ),
    Pendientes AS
    (
        SELECT
            p.Id,
            COALESCE(m.UltimoNumero, 0) + ROW_NUMBER() OVER (PARTITION BY p.NegocioId ORDER BY p.Id) AS NumeroPorNegocio
        FROM dbo.PromocionesHorario p
        LEFT JOIN Maximos m ON m.NegocioId = p.NegocioId
        WHERE p.NumeroPorNegocio IS NULL OR p.NumeroPorNegocio <= 0
    )
    UPDATE p
    SET NumeroPorNegocio = CONVERT(INT, x.NumeroPorNegocio)
    FROM dbo.PromocionesHorario p
    INNER JOIN Pendientes x ON x.Id = p.Id;

    DECLARE @Maximos TABLE
    (
        NegocioId INT NOT NULL,
        Entidad NVARCHAR(30) NOT NULL,
        UltimoNumero INT NOT NULL,
        PRIMARY KEY (NegocioId, Entidad)
    );

    INSERT INTO @Maximos (NegocioId, Entidad, UltimoNumero)
    SELECT s.NegocioId, N'SEDE', MAX(s.NumeroPorNegocio)
    FROM dbo.Sedes s
    GROUP BY s.NegocioId
    UNION ALL
    SELECT s.NegocioId, N'ESPACIO', MAX(e.NumeroPorNegocio)
    FROM dbo.EspaciosDeportivos e
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    GROUP BY s.NegocioId
    UNION ALL
    SELECT c.NegocioId, N'CLIENTE', MAX(c.NumeroPorNegocio)
    FROM dbo.Clientes c
    GROUP BY c.NegocioId
    UNION ALL
    SELECT s.NegocioId, N'RESERVA', MAX(r.NumeroPorNegocio)
    FROM dbo.Reservas r
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    GROUP BY s.NegocioId
    UNION ALL
    SELECT s.NegocioId, N'PAGO', MAX(p.NumeroPorNegocio)
    FROM dbo.Pagos p
    INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
    INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
    INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
    GROUP BY s.NegocioId
    UNION ALL
    SELECT c.NegocioId, N'CUPON', MAX(c.NumeroPorNegocio)
    FROM dbo.Cupones c
    GROUP BY c.NegocioId
    UNION ALL
    SELECT p.NegocioId, N'PROMOCION', MAX(p.NumeroPorNegocio)
    FROM dbo.PromocionesHorario p
    GROUP BY p.NegocioId;

    UPDATE nc
    SET UltimoNumero = CASE WHEN nc.UltimoNumero < m.UltimoNumero THEN m.UltimoNumero ELSE nc.UltimoNumero END,
        FechaActualizacion = SYSUTCDATETIME(),
        UsuarioActualizacion = N'MIGRACION_20261008'
    FROM dbo.NegocioCorrelativos nc
    INNER JOIN @Maximos m
        ON m.NegocioId = nc.NegocioId
       AND m.Entidad = nc.Entidad;

    INSERT INTO dbo.NegocioCorrelativos
    (
        NegocioId,
        Entidad,
        UltimoNumero,
        FechaActualizacion,
        UsuarioActualizacion
    )
    SELECT
        m.NegocioId,
        m.Entidad,
        m.UltimoNumero,
        SYSUTCDATETIME(),
        N'MIGRACION_20261008'
    FROM @Maximos m
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM dbo.NegocioCorrelativos nc WITH (UPDLOCK, HOLDLOCK)
        WHERE nc.NegocioId = m.NegocioId
          AND nc.Entidad = m.Entidad
    );

    IF EXISTS (SELECT 1 FROM dbo.Sedes WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.EspaciosDeportivos WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.Clientes WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.Reservas WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.Pagos WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.Cupones WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
       OR EXISTS (SELECT 1 FROM dbo.PromocionesHorario WHERE NumeroPorNegocio IS NULL OR NumeroPorNegocio <= 0)
        RAISERROR('No fue posible completar todos los correlativos historicos.', 16, 1);

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
GO
