-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/07/2026
-- Description:   Permite lineas analiticas de ajuste cambiario en CON_AsientoDetalle con Debe/Haber en cero y saldo conservado en TotalImporteS o TotalImporteD.
-- =============================================
-- Firma: FRANCO LARA - 06/07/2026 | Ajusta la restriccion CK_CON_AsientoDetalle_Montos para aceptar lineas analiticas de cancelacion total con Debe/Haber en cero y equivalencia pendiente por moneda.

IF OBJECT_ID(N'dbo.CON_AsientoDetalle', N'U') IS NOT NULL
BEGIN
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
                    OR (
                        Debe = 0
                        AND Haber = 0
                        AND (
                            TotalImporteS > 0
                            OR TotalImporteD > 0
                        )
                    )
                )
            );
END;
