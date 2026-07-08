-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/07/2026
-- Description:   Agrega la columna DH al detalle contable y actualiza la restriccion de montos para guardar explicitamente el sentido Debe/Haber.
-- =============================================

IF OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH(N'dbo.CON_AsientoDetalle', N'DH') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            ADD DH CHAR(1) NULL;
    END;

    UPDATE d
    SET DH = CASE
                 WHEN d.Debe > 0 THEN 'D'
                 WHEN d.Haber > 0 THEN 'H'
                 ELSE ISNULL(d.DH, 'D')
             END
    FROM dbo.CON_AsientoDetalle AS d
    WHERE d.DH IS NULL
       OR d.DH NOT IN ('D', 'H');

    IF EXISTS
    (
        SELECT 1
        FROM sys.columns
        WHERE object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
          AND name = N'DH'
          AND is_nullable = 1
    )
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            ALTER COLUMN DH CHAR(1) NOT NULL;
    END;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.default_constraints
        WHERE parent_object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
          AND name = N'DF_CON_AsientoDetalle_DH'
    )
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            ADD CONSTRAINT DF_CON_AsientoDetalle_DH DEFAULT ('D') FOR DH;
    END;

    IF EXISTS
    (
        SELECT 1
        FROM sys.check_constraints
        WHERE parent_object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle')
          AND name = N'CK_CON_AsientoDetalle_Montos'
    )
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            DROP CONSTRAINT CK_CON_AsientoDetalle_Montos;
    END;

    ALTER TABLE dbo.CON_AsientoDetalle
        ADD CONSTRAINT CK_CON_AsientoDetalle_Montos
            CHECK (
                DH IN ('D', 'H')
                AND Debe >= 0
                AND Haber >= 0
                AND (
                    (DH = 'D' AND Debe > 0 AND Haber = 0)
                    OR (DH = 'H' AND Debe = 0 AND Haber > 0)
                )
            );
END;
