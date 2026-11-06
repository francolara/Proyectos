-- Firma: FRANCO LARA | Create date: 06/11/2026 | Numeración visible del presupuesto por año/mes de registro en Lima y correlativo mensual.
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF COL_LENGTH(N'dbo.Presupuesto',N'NumeroPresupuesto') IS NULL ALTER TABLE dbo.Presupuesto ADD NumeroPresupuesto NVARCHAR(40) NULL;
GO
BEGIN TRY
    BEGIN TRANSACTION;
    DECLARE @Filas BIGINT;
    SELECT @Filas=COUNT_BIG(1) FROM dbo.Presupuesto WITH(TABLOCKX,HOLDLOCK);
    ;WITH Numerados AS
    (
        SELECT p.PresupuestoId,CONVERT(CHAR(6),CONVERT(DATETIME2,p.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112) AS Periodo,
            ROW_NUMBER() OVER(PARTITION BY CONVERT(CHAR(6),CONVERT(DATETIME2,p.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time'),112) ORDER BY p.FechaRegistro,p.PresupuestoId) AS Secuencia
        FROM dbo.Presupuesto AS p
    )
    UPDATE p SET NumeroPresupuesto=CONCAT(n.Periodo,CASE WHEN n.Secuencia<1000000 THEN RIGHT(CONCAT(N'000000',n.Secuencia),6) ELSE CONVERT(NVARCHAR(20),n.Secuencia) END)
    FROM dbo.Presupuesto AS p INNER JOIN Numerados AS n ON n.PresupuestoId=p.PresupuestoId WHERE p.NumeroPresupuesto IS NULL;
    ;WITH Maximos AS(SELECT LEFT(p.NumeroPresupuesto,6) AS Periodo,MAX(TRY_CONVERT(BIGINT,SUBSTRING(p.NumeroPresupuesto,7,34))) AS Ultimo FROM dbo.Presupuesto AS p GROUP BY LEFT(p.NumeroPresupuesto,6))
    UPDATE contador SET Ultimo=maximo.Ultimo FROM dbo.PresupuestoCorrelativoMensual AS contador INNER JOIN Maximos AS maximo ON maximo.Periodo=contador.Periodo WHERE contador.Ultimo<maximo.Ultimo;
    INSERT INTO dbo.PresupuestoCorrelativoMensual(Periodo,Ultimo)
    SELECT LEFT(p.NumeroPresupuesto,6),MAX(TRY_CONVERT(BIGINT,SUBSTRING(p.NumeroPresupuesto,7,34))) FROM dbo.Presupuesto AS p
    WHERE NOT EXISTS(SELECT 1 FROM dbo.PresupuestoCorrelativoMensual AS c WHERE c.Periodo=LEFT(p.NumeroPresupuesto,6)) GROUP BY LEFT(p.NumeroPresupuesto,6);
    ALTER TABLE dbo.Presupuesto ALTER COLUMN NumeroPresupuesto NVARCHAR(40) NOT NULL;
    IF NOT EXISTS(SELECT 1 FROM sys.indexes AS i WHERE i.object_id=OBJECT_ID(N'dbo.Presupuesto') AND i.name=N'UX_Presupuesto_NumeroPresupuesto') CREATE UNIQUE INDEX UX_Presupuesto_NumeroPresupuesto ON dbo.Presupuesto(NumeroPresupuesto);
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
    DECLARE @Mensaje NVARCHAR(4000)=ERROR_MESSAGE(),@Severidad INT=ERROR_SEVERITY(),@Estado INT=ERROR_STATE();
    RAISERROR(@Mensaje,@Severidad,@Estado);
END CATCH;
GO
