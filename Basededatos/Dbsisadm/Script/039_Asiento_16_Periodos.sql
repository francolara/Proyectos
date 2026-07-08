-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Amplia el modulo de asientos a 16 periodos contables y ajusta sus restricciones para admitir meses 00-15 en cabecera y correlativos.
-- =============================================

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_CON_Asiento_Mes'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_Asiento')
)
BEGIN
    ALTER TABLE dbo.CON_Asiento
        DROP CONSTRAINT CK_CON_Asiento_Mes;
END;

ALTER TABLE dbo.CON_Asiento
    ADD CONSTRAINT CK_CON_Asiento_Mes
        CHECK (Mes BETWEEN 0 AND 15);

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_CON_Asiento_Periodo'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_Asiento')
)
BEGIN
    ALTER TABLE dbo.CON_Asiento
        DROP CONSTRAINT CK_CON_Asiento_Periodo;
END;

ALTER TABLE dbo.CON_Asiento
    ADD CONSTRAINT CK_CON_Asiento_Periodo
        CHECK (
            Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
            AND Periodo = CONVERT(CHAR(4), Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), Mes), 2)
        );

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_CON_CorrelativoAsiento_Periodo'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_CorrelativoAsiento')
)
BEGIN
    ALTER TABLE dbo.CON_CorrelativoAsiento
        DROP CONSTRAINT CK_CON_CorrelativoAsiento_Periodo;
END;

ALTER TABLE dbo.CON_CorrelativoAsiento
    ADD CONSTRAINT CK_CON_CorrelativoAsiento_Periodo
        CHECK (
            Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
            AND RIGHT(Periodo, 2) BETWEEN '00' AND '15'
        );
