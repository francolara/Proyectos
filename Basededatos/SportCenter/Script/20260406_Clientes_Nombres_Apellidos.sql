-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Agrega Nombres y Apellidos en clientes y rellena NombresORazonSocial para compatibilidad.
-- =============================================

IF COL_LENGTH('dbo.Clientes', 'Nombres') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD Nombres NVARCHAR(120) NULL;
END;

IF COL_LENGTH('dbo.Clientes', 'Apellidos') IS NULL
BEGIN
    ALTER TABLE dbo.Clientes ADD Apellidos NVARCHAR(120) NULL;
END;

UPDATE c
SET c.Nombres = CASE WHEN c.TipoDocumento <> N'6' AND c.Nombres IS NULL THEN NULLIF(LTRIM(RTRIM(c.NombresORazonSocial)), N'') ELSE c.Nombres END,
    c.Apellidos = CASE WHEN c.TipoDocumento = N'6' THEN NULL ELSE c.Apellidos END,
    c.NombresORazonSocial = CASE
        WHEN c.TipoDocumento = N'6' THEN LTRIM(RTRIM(c.NombresORazonSocial))
        ELSE LEFT(LTRIM(RTRIM(CONCAT(COALESCE(c.Nombres, N''), N' ', COALESCE(c.Apellidos, N'')))), 200)
    END
FROM dbo.Clientes c;