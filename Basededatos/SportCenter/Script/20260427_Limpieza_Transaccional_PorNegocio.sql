USE [DbSportCenter];
GO

SET ANSI_NULLS ON;
GO
SET QUOTED_IDENTIFIER ON;
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- =============================================
-- Firma: Codex - 27/04/2026 | Script de limpieza transaccional por NegocioId con modo PREVIEW/EXECUTE, borrado en orden de dependencias FK y resumen de filas afectadas.

/*
    OBJETIVO
    - Limpiar SOLO tablas transaccionales del negocio indicado.
    - No elimina configuraciones, maestros ni estructura del negocio.

    USO
    1) Simular (recomendado):
       - Dejar @Ejecutar = 0 y revisar conteos.
    2) Ejecutar:
       - Cambiar @Ejecutar = 1.

    TABLAS TRANSACCIONALES INCLUIDAS
    - dbo.ComprobantesDetalle
    - dbo.ComprobantesElectronicos
    - dbo.Pagos
    - dbo.SolicitudesReservaPublica
    - dbo.ReservasUsuariosPublicos
    - dbo.Reservas
    - dbo.BloqueosHorario
    - dbo.NegocioNotificaciones
    - dbo.BitacoraAuditoria (solo del negocio)
*/

DECLARE @NegocioId INT = 1;      -- TODO: Reemplazar por el negocio objetivo
DECLARE @Ejecutar BIT = 0;       -- 0 = PREVIEW | 1 = EXECUTE

IF @NegocioId IS NULL OR @NegocioId <= 0
BEGIN
    RAISERROR('Debes indicar un @NegocioId valido.', 16, 1);
    RETURN;
END;

IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
BEGIN
    RAISERROR('No existe el NegocioId indicado.', 16, 1);
    RETURN;
END;

DECLARE @Sedes TABLE (Id INT PRIMARY KEY);
DECLARE @Espacios TABLE (Id INT PRIMARY KEY);
DECLARE @Clientes TABLE (Id INT PRIMARY KEY);
DECLARE @Reservas TABLE (Id INT PRIMARY KEY);
DECLARE @Comprobantes TABLE (Id INT PRIMARY KEY);

INSERT INTO @Sedes (Id)
SELECT s.Id
FROM dbo.Sedes s
WHERE s.NegocioId = @NegocioId;

INSERT INTO @Espacios (Id)
SELECT e.Id
FROM dbo.EspaciosDeportivos e
INNER JOIN @Sedes s ON s.Id = e.SedeId;

INSERT INTO @Clientes (Id)
SELECT c.Id
FROM dbo.Clientes c
WHERE c.NegocioId = @NegocioId;

INSERT INTO @Reservas (Id)
SELECT DISTINCT r.Id
FROM dbo.Reservas r
LEFT JOIN @Espacios e ON e.Id = r.EspacioDeportivoId
LEFT JOIN @Clientes c ON c.Id = r.ClienteId
WHERE e.Id IS NOT NULL
   OR c.Id IS NOT NULL;

INSERT INTO @Comprobantes (Id)
SELECT ce.Id
FROM dbo.ComprobantesElectronicos ce
WHERE ce.NegocioId = @NegocioId
   OR EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = ce.ReservaId);

DECLARE @Resumen TABLE
(
    Tabla NVARCHAR(128) NOT NULL,
    Filas INT NOT NULL
);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'ComprobantesDetalle', COUNT(1)
FROM dbo.ComprobantesDetalle cd
WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = cd.ComprobanteElectronicoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'ComprobantesElectronicos', COUNT(1)
FROM dbo.ComprobantesElectronicos ce
WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = ce.Id);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'Pagos', COUNT(1)
FROM dbo.Pagos p
WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = p.ReservaId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'SolicitudesReservaPublica', COUNT(1)
FROM dbo.SolicitudesReservaPublica srp
WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = srp.EspacioDeportivoId)
   OR EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = srp.ReservaId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'ReservasUsuariosPublicos', COUNT(1)
FROM dbo.ReservasUsuariosPublicos rup
WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = rup.ReservaId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'Reservas', COUNT(1)
FROM dbo.Reservas r
WHERE EXISTS (SELECT 1 FROM @Reservas rx WHERE rx.Id = r.Id);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'BloqueosHorario', COUNT(1)
FROM dbo.BloqueosHorario bh
WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = bh.EspacioDeportivoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'NegocioNotificaciones', COUNT(1)
FROM dbo.NegocioNotificaciones nn
WHERE nn.NegocioId = @NegocioId;

INSERT INTO @Resumen (Tabla, Filas)
SELECT 'BitacoraAuditoria', COUNT(1)
FROM dbo.BitacoraAuditoria ba
WHERE ba.NegocioId = @NegocioId;

PRINT '=== PREVIEW LIMPIEZA TRANSACCIONAL ===';
SELECT @NegocioId AS NegocioId, n.NombreORazonSocial AS Negocio
FROM dbo.Negocios n
WHERE n.Id = @NegocioId;

SELECT Tabla, Filas
FROM @Resumen
ORDER BY Tabla;

IF @Ejecutar = 0
BEGIN
    PRINT 'Modo PREVIEW activo. No se elimino informacion.';
    RETURN;
END;

BEGIN TRY
    BEGIN TRAN;

    DELETE cd
    FROM dbo.ComprobantesDetalle cd
    WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = cd.ComprobanteElectronicoId);

    DELETE ce
    FROM dbo.ComprobantesElectronicos ce
    WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = ce.Id);

    DELETE p
    FROM dbo.Pagos p
    WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = p.ReservaId);

    DELETE srp
    FROM dbo.SolicitudesReservaPublica srp
    WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = srp.EspacioDeportivoId)
       OR EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = srp.ReservaId);

    DELETE rup
    FROM dbo.ReservasUsuariosPublicos rup
    WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = rup.ReservaId);

    DELETE r
    FROM dbo.Reservas r
    WHERE EXISTS (SELECT 1 FROM @Reservas rx WHERE rx.Id = r.Id);

    DELETE bh
    FROM dbo.BloqueosHorario bh
    WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = bh.EspacioDeportivoId);

    DELETE nn
    FROM dbo.NegocioNotificaciones nn
    WHERE nn.NegocioId = @NegocioId;

    DELETE ba
    FROM dbo.BitacoraAuditoria ba
    WHERE ba.NegocioId = @NegocioId;

    COMMIT;
    PRINT 'Limpieza transaccional ejecutada correctamente.';
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;

