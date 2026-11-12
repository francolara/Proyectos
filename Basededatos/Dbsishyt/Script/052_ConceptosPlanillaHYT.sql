-- Firma: FRANCO LARA | Create date: 10/11/2026 | Siembra conceptos habituales con identidad estable para asignación familiar y autoriza el acceso al maestro.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
INSERT INTO dbo.ConceptoPlanillaHYT(Nombre,Naturaleza,Calculo,Regimen,CodigoSistema,Activo)
SELECT semilla.Nombre,semilla.Naturaleza,semilla.Calculo,semilla.Regimen,semilla.CodigoSistema,1
FROM (VALUES
    (N'Asignación familiar','ABONO','IMPORTE',NULL,'ASIGNACION_FAMILIAR'),
    (N'Bono','ABONO','IMPORTE',NULL,NULL),
    (N'Sistema de pensiones (AFP)','CARGO','PORCENTAJE','AFP',NULL),
    (N'Sistema de pensiones (ONP)','CARGO','PORCENTAJE','ONP',NULL)
) AS semilla(Nombre,Naturaleza,Calculo,Regimen,CodigoSistema)
WHERE NOT EXISTS (SELECT 1 FROM dbo.ConceptoPlanillaHYT AS concepto
    WHERE concepto.Nombre=semilla.Nombre OR concepto.Regimen=semilla.Regimen OR concepto.CodigoSistema=semilla.CodigoSistema);
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ConceptoPlanillaHYT_Listar TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ConceptoPlanillaHYT_Guardar TO hyt_app_executor;
GO
