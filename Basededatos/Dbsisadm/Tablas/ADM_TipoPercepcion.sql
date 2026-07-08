-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Maestro general de tipos de percepcion aplicables a compras.
-- =============================================

IF OBJECT_ID(N'dbo.ADM_TipoPercepcion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ADM_TipoPercepcion
    (
        IdTipoPercepcion INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_ADM_TipoPercepcion PRIMARY KEY,
        Codigo VARCHAR(2) NOT NULL,
        Descripcion NVARCHAR(200) NOT NULL,
        Porcentaje DECIMAL(7,4) NOT NULL CONSTRAINT DF_ADM_TipoPercepcion_Porcentaje DEFAULT (0),
        Estado BIT NOT NULL CONSTRAINT DF_ADM_TipoPercepcion_Estado DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_ADM_TipoPercepcion_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.ADM_TipoPercepcion
        ADD CONSTRAINT UQ_ADM_TipoPercepcion_Codigo UNIQUE (Codigo);

    ALTER TABLE dbo.ADM_TipoPercepcion
        ADD CONSTRAINT CK_ADM_TipoPercepcion_Porcentaje CHECK (Porcentaje >= 0);
END;
