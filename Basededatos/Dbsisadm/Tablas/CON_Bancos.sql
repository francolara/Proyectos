-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Crea el catalogo maestro de bancos y precarga entidades financieras para ayuda de cuentas corrientes.
-- =============================================

IF OBJECT_ID(N'dbo.CON_Bancos', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_Bancos
    (
        IdBanco INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_Bancos PRIMARY KEY,
        Codigo VARCHAR(20) NOT NULL,
        Nombre NVARCHAR(200) NOT NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_Bancos_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_Bancos
        ADD CONSTRAINT UQ_CON_Bancos_Codigo UNIQUE (Codigo);
END;

;WITH BancosBase AS
(
    SELECT *
    FROM (VALUES
        ('BCP',        N'Banco de Credito del Peru - BCP', 1),
        ('INTERBANK',  N'Interbank', 1),
        ('BBVA',       N'BBVA Peru', 1),
        ('SCOTIA',     N'Scotiabank Peru', 1),
        ('NACION',     N'Banco de la Nacion', 1),
        ('BANBIF',     N'BanBif - Banco Interamericano de Finanzas', 1),
        ('PICHINCHA',  N'Banco Pichincha', 1),
        ('MIBANCO',    N'Mibanco', 1),
        ('FALABELLA',  N'Banco Falabella', 1),
        ('RIPLEY',     N'Banco Ripley', 1),
        ('CITIBANK',   N'Citibank del Peru', 1),
        ('SANTANDER',  N'Banco Santander Peru', 1),
        ('ALFIN',      N'Alfin Banco', 1),
        ('COMERCIO',   N'Banco de Comercio', 1),
        ('BCI',        N'Banco BCI Peru', 1),
        ('GNB',        N'Banco GNB Peru', 1),
        ('ICBC',       N'Banco ICBC Peru', 1),
        ('EFECTIVA',   N'Banco Efectiva', 1),
        ('AZTECA',     N'Banco Azteca Peru', 1),
        ('BANCOM',     N'Bancom', 1)
    ) AS fuente (Codigo, Nombre, Estado)
)
MERGE dbo.CON_Bancos AS destino
USING BancosBase AS fuente
    ON destino.Codigo = fuente.Codigo
WHEN MATCHED THEN
    UPDATE SET
        destino.Nombre = fuente.Nombre,
        destino.Estado = fuente.Estado
WHEN NOT MATCHED BY TARGET THEN
    INSERT (Codigo, Nombre, Estado)
    VALUES (fuente.Codigo, fuente.Nombre, fuente.Estado);
