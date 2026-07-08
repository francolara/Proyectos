-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Historial de generación de libros electrónicos PLE por empresa, periodo y formato.
-- =============================================

IF OBJECT_ID(N'dbo.CON_LibroElectronicoGeneracion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_LibroElectronicoGeneracion
    (
        IdLibroElectronicoGeneracion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_LibroElectronicoGeneracion PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        Periodo CHAR(6) NOT NULL,
        CodigoLibro VARCHAR(10) NOT NULL,
        CodigoFormato VARCHAR(10) NOT NULL,
        NombreArchivo NVARCHAR(250) NOT NULL,
        CantidadRegistros INT NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_CantidadRegistros DEFAULT (0),
        TotalDebe DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_TotalDebe DEFAULT (0),
        TotalHaber DECIMAL(18,2) NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_TotalHaber DEFAULT (0),
        Estado NVARCHAR(20) NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_Estado DEFAULT (N'GENERADO'),
        Observaciones NVARCHAR(MAX) NULL,
        FechaGeneracion DATETIME2(0) NOT NULL CONSTRAINT DF_CON_LibroElectronicoGeneracion_FechaGeneracion DEFAULT (SYSDATETIME()),
        UsuarioGeneracion NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_LibroElectronicoGeneracion
        ADD CONSTRAINT FK_CON_LibroElectronicoGeneracion_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_LibroElectronicoGeneracion
        ADD CONSTRAINT CK_CON_LibroElectronicoGeneracion_Periodo
            CHECK (Periodo LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]');

    ALTER TABLE dbo.CON_LibroElectronicoGeneracion
        ADD CONSTRAINT CK_CON_LibroElectronicoGeneracion_Totales
            CHECK (CantidadRegistros >= 0 AND TotalDebe >= 0 AND TotalHaber >= 0);

    CREATE INDEX IX_CON_LibroElectronicoGeneracion_EmpresaPeriodo
        ON dbo.CON_LibroElectronicoGeneracion (IdEmpresa, Periodo, CodigoLibro, FechaGeneracion DESC);
END;
