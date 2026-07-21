-- =============================================
-- Author:        FRANCO LARA
-- Create date:   21/07/2026
-- Firma:         Persiste el plan comercial vigente, agrega trazabilidad por cobro y publica la fecha de registro para reportes HTML.
-- Base destino:  dbsportcenter_20260613
-- =============================================

IF COL_LENGTH(N'dbo.NegociosSuscripcion', N'PlanComercial') IS NULL
BEGIN
    ALTER TABLE dbo.NegociosSuscripcion
        ADD PlanComercial NVARCHAR(20) NULL;
END;

UPDATE ns
SET PlanComercial = CASE
    WHEN COALESCE(ns.EsPrueba, 0) = 1 THEN N'PRUEBA'
    WHEN UPPER(LTRIM(RTRIM(COALESCE(n.TipoPlan, N'')))) = N'FULL' THEN N'PRO'
    ELSE N'ESENCIAL'
END
FROM dbo.NegociosSuscripcion ns
INNER JOIN dbo.Negocios n ON n.Id = ns.NegocioId
WHERE NULLIF(LTRIM(RTRIM(ns.PlanComercial)), N'') IS NULL;

ALTER TABLE dbo.NegociosSuscripcion ALTER COLUMN PlanComercial NVARCHAR(20) NOT NULL;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints dc
    INNER JOIN sys.columns c ON c.object_id = dc.parent_object_id AND c.column_id = dc.parent_column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'dbo.NegociosSuscripcion')
      AND c.name = N'PlanComercial'
)
BEGIN
    ALTER TABLE dbo.NegociosSuscripcion
        ADD CONSTRAINT DF_NegociosSuscripcion_PlanComercial DEFAULT (N'PRUEBA') FOR PlanComercial;
END;

IF NOT EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.NegociosSuscripcion')
      AND name = N'CK_NegociosSuscripcion_PlanComercial'
)
BEGIN
    ALTER TABLE dbo.NegociosSuscripcion WITH CHECK
        ADD CONSTRAINT CK_NegociosSuscripcion_PlanComercial
            CHECK (PlanComercial IN (N'PRUEBA', N'ESENCIAL', N'PRO'));
END;

IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'PlanComercialObjetivo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionPago ADD PlanComercialObjetivo NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'TipoPlanObjetivo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionPago ADD TipoPlanObjetivo NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'SedesPermitidasObjetivo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionPago ADD SedesPermitidasObjetivo INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'EspaciosPermitidosObjetivo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionPago ADD EspaciosPermitidosObjetivo INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionPago', N'UsuariosPermitidosObjetivo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionPago ADD UsuariosPermitidosObjetivo INT NULL;

IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'PlanComercialAnterior') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD PlanComercialAnterior NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'PlanComercialNuevo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD PlanComercialNuevo NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'TipoPlanAnterior') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD TipoPlanAnterior NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'TipoPlanNuevo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD TipoPlanNuevo NVARCHAR(20) NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'SedesPermitidasAnterior') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD SedesPermitidasAnterior INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'SedesPermitidasNuevo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD SedesPermitidasNuevo INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'EspaciosPermitidosAnterior') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD EspaciosPermitidosAnterior INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'EspaciosPermitidosNuevo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD EspaciosPermitidosNuevo INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'UsuariosPermitidosAnterior') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD UsuariosPermitidosAnterior INT NULL;
IF COL_LENGTH(N'dbo.NegociosSuscripcionMovimiento', N'UsuariosPermitidosNuevo') IS NULL
    ALTER TABLE dbo.NegociosSuscripcionMovimiento ADD UsuariosPermitidosNuevo INT NULL;

-- Despues de este script, publicar los procedimientos almacenados modificados incluidos en StoreProcedure.
-- Sp_Plataforma_Negocios_Listar incluye n.FechaRegistro como ultima columna para los reportes del Super Admin.
