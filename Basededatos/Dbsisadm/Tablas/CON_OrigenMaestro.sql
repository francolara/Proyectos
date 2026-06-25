-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Catalogo maestro interno de origenes contables base. No pertenece a una empresa.
-- =============================================

IF OBJECT_ID(N'dbo.CON_OrigenMaestro', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_OrigenMaestro
    (
        IdOrigenMaestro INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_OrigenMaestro PRIMARY KEY,
        CodigoOrigen VARCHAR(10) NOT NULL,
        NombreOrigen NVARCHAR(150) NOT NULL,
        ModuloOrigen NVARCHAR(50) NOT NULL,
        PermiteRegistroManual BIT NOT NULL CONSTRAINT DF_CON_OrigenMaestro_PermiteRegistroManual DEFAULT (1),
        Estado BIT NOT NULL CONSTRAINT DF_CON_OrigenMaestro_Estado DEFAULT (1),
        Orden INT NOT NULL CONSTRAINT DF_CON_OrigenMaestro_Orden DEFAULT (0),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_OrigenMaestro_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_OrigenMaestro
        ADD CONSTRAINT UQ_CON_OrigenMaestro_CodigoOrigen
            UNIQUE (CodigoOrigen);
END;
