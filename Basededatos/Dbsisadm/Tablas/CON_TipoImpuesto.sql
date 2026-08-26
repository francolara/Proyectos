-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Maestro interno de tipos de impuesto usados en configuracion contable.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Agrega cuenta contable maestra unica para impuesto.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/06/2026
-- Description:   Incorpora el tipo de impuesto SPOT para configurar la cuenta contable de detracciones.
-- =============================================

-- Firma: FRANCO LARA - 25/08/2026 | Reemplaza el identificador de plan contable del maestro por el codigo de cuenta portable entre empresas.

IF OBJECT_ID(N'dbo.CON_TipoImpuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_TipoImpuesto
    (
        IdTipoImpuesto INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_TipoImpuesto PRIMARY KEY,
        CodigoSunat VARCHAR(10) NOT NULL,
        NombreImpuesto NVARCHAR(100) NOT NULL,
        CodigoCuenta VARCHAR(20) NULL,
        Estado BIT NOT NULL CONSTRAINT DF_CON_TipoImpuesto_Estado DEFAULT (1)
    );

    ALTER TABLE dbo.CON_TipoImpuesto
        ADD CONSTRAINT UQ_CON_TipoImpuesto_CodigoSunat UNIQUE (CodigoSunat);
END;
