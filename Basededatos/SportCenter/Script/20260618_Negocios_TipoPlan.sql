-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Firma:         Agrega TipoPlan en Negocios con default Basico y normaliza valores existentes para control comercial por plan.
-- =============================================

IF COL_LENGTH('dbo.Negocios', 'TipoPlan') IS NULL
BEGIN
    ALTER TABLE dbo.Negocios
    ADD TipoPlan NVARCHAR(20) NOT NULL
        CONSTRAINT DF_Negocios_TipoPlan DEFAULT (N'Basico');
END;

UPDATE dbo.Negocios
SET TipoPlan = N'Basico'
WHERE TipoPlan IS NULL
   OR LTRIM(RTRIM(TipoPlan)) = N'';

UPDATE dbo.Negocios
SET TipoPlan = CASE
                   WHEN UPPER(LTRIM(RTRIM(TipoPlan))) = N'FULL' THEN N'Full'
                   ELSE N'Basico'
               END
WHERE TipoPlan IS NOT NULL;

IF NOT EXISTS (
    SELECT 1
    FROM sys.check_constraints
    WHERE name = N'CK_Negocios_TipoPlan'
      AND parent_object_id = OBJECT_ID(N'dbo.Negocios')
)
BEGIN
    ALTER TABLE dbo.Negocios WITH CHECK
    ADD CONSTRAINT CK_Negocios_TipoPlan
        CHECK (TipoPlan IN (N'Basico', N'Full'));
END;
