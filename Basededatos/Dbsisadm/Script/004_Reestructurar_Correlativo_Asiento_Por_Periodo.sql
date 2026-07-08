-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Ajusta la cabecera contable para correlativo por empresa, origen y periodo; ademas inicializa la tabla de numeradores.
-- =============================================
-- Firma: FRANCO LARA - 02/07/2026 | Adecua las restricciones de asiento y correlativo para admitir periodos contables 00-15 sin perder el correlativo por periodo.

IF COL_LENGTH(N'dbo.CON_Asiento', N'Periodo') IS NULL
BEGIN
    ALTER TABLE dbo.CON_Asiento
        ADD Periodo CHAR(6) NULL;
END;

IF COL_LENGTH(N'dbo.CON_Asiento', N'Periodo') IS NOT NULL
BEGIN
    UPDATE dbo.CON_Asiento
    SET Periodo = CONVERT(CHAR(4), Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), Mes), 2)
    WHERE Periodo IS NULL
       OR Periodo <> CONVERT(CHAR(4), Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), Mes), 2);

    IF NOT EXISTS
    (
        SELECT 1
        FROM dbo.CON_Asiento
        WHERE Periodo IS NULL
    )
    BEGIN
        ALTER TABLE dbo.CON_Asiento
            ALTER COLUMN Periodo CHAR(6) NOT NULL;
    END;
END;

IF EXISTS
(
    SELECT 1
    FROM sys.key_constraints
    WHERE name = N'UQ_CON_Asiento_Numero'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_Asiento')
)
BEGIN
    ALTER TABLE dbo.CON_Asiento
        DROP CONSTRAINT UQ_CON_Asiento_Numero;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_CON_Asiento_Periodo'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_Asiento')
)
BEGIN
    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT CK_CON_Asiento_Periodo
            CHECK (
                Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND Mes BETWEEN 0 AND 15
                AND Periodo = CONVERT(CHAR(4), Ejercicio) + RIGHT('0' + CONVERT(VARCHAR(2), Mes), 2)
            );
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.key_constraints
    WHERE name = N'UQ_CON_Asiento_Numero'
      AND parent_object_id = OBJECT_ID(N'dbo.CON_Asiento')
)
BEGIN
    ALTER TABLE dbo.CON_Asiento
        ADD CONSTRAINT UQ_CON_Asiento_Numero
            UNIQUE (IdEmpresa, IdOrigen, Periodo, NumeroAsiento);
END;

IF OBJECT_ID(N'dbo.CON_CorrelativoAsiento', N'U') IS NOT NULL
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
    SELECT
        a.IdEmpresa,
        a.IdOrigen,
        a.Periodo,
        MAX(a.NumeroAsiento) AS UltimoNumero,
        SYSDATETIME(),
        MAX(a.UsuarioRegistro) AS UsuarioRegistro
    FROM dbo.CON_Asiento AS a
    GROUP BY
        a.IdEmpresa,
        a.IdOrigen,
        a.Periodo
    HAVING NOT EXISTS
    (
        SELECT 1
        FROM dbo.CON_CorrelativoAsiento AS c
        WHERE c.IdEmpresa = a.IdEmpresa
          AND c.IdOrigen = a.IdOrigen
          AND c.Periodo = a.Periodo
    );
END;
