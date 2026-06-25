-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Amplia TipoDocumento del detalle de asiento para guardar la descripcion del comprobante en compras y ventas.
-- =============================================

IF OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U') IS NOT NULL
BEGIN
    IF COL_LENGTH('dbo.CON_AsientoDetalle', 'TipoDocumento') IS NULL
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            ADD TipoDocumento NVARCHAR(150) NULL;
    END;
    ELSE IF EXISTS
    (
        SELECT 1
        FROM sys.columns AS c
        INNER JOIN sys.types AS t
            ON t.user_type_id = c.user_type_id
        WHERE c.object_id = OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U')
          AND c.name = 'TipoDocumento'
          AND (
                t.name <> 'nvarchar'
                OR c.max_length < 300
              )
    )
    BEGIN
        ALTER TABLE dbo.CON_AsientoDetalle
            ALTER COLUMN TipoDocumento NVARCHAR(150) NULL;
    END;
END;
