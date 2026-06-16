-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Numerador contable por empresa, origen y periodo para reinicio mensual de comprobantes.
-- =============================================

IF OBJECT_ID(N'dbo.CON_CorrelativoAsiento', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_CorrelativoAsiento
    (
        IdCorrelativoAsiento INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_CorrelativoAsiento PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        IdOrigen INT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        UltimoNumero INT NOT NULL CONSTRAINT DF_CON_CorrelativoAsiento_UltimoNumero DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_CorrelativoAsiento_FechaRegistro DEFAULT (SYSDATETIME()),
        FechaActualizacion DATETIME2(0) NULL,
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_CorrelativoAsiento
        ADD CONSTRAINT FK_CON_CorrelativoAsiento_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_CorrelativoAsiento
        ADD CONSTRAINT FK_CON_CorrelativoAsiento_CON_Origen
            FOREIGN KEY (IdOrigen) REFERENCES dbo.CON_Origen (IdOrigen);

    ALTER TABLE dbo.CON_CorrelativoAsiento
        ADD CONSTRAINT CK_CON_CorrelativoAsiento_Periodo
            CHECK (
                Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
                AND RIGHT(Periodo, 2) BETWEEN '01' AND '12'
            );

    ALTER TABLE dbo.CON_CorrelativoAsiento
        ADD CONSTRAINT CK_CON_CorrelativoAsiento_UltimoNumero
            CHECK (UltimoNumero >= 0);

    ALTER TABLE dbo.CON_CorrelativoAsiento
        ADD CONSTRAINT UQ_CON_CorrelativoAsiento
            UNIQUE (IdEmpresa, IdOrigen, Periodo);
END;
