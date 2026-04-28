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
-- Firma: Codex - 27/04/2026 | Script de limpieza total (negocios, sedes, clientes, reservas y usuarios) conservando solo cuentas con rol OwnerPlataforma, con modo PREVIEW/EXECUTE y borrado seguro por dependencias FK.

/*
    OBJETIVO
    - Eliminar informacion operativa completa de la plataforma:
      negocios, sedes, espacios, clientes, reservas y usuarios.
    - Conservar unicamente los usuarios con rol OwnerPlataforma (superadmin).

    USO
    1) PREVIEW (recomendado): @Ejecutar = 0
    2) EXECUTE: @Ejecutar = 1
*/

DECLARE @Ejecutar BIT = 0; -- 0 = PREVIEW | 1 = EXECUTE

DECLARE @SuperAdmins TABLE (UsuarioId NVARCHAR(450) PRIMARY KEY);
DECLARE @UsuariosEliminar TABLE (UsuarioId NVARCHAR(450) PRIMARY KEY);
DECLARE @Negocios TABLE (Id INT PRIMARY KEY);
DECLARE @Sedes TABLE (Id INT PRIMARY KEY);
DECLARE @Espacios TABLE (Id INT PRIMARY KEY);
DECLARE @Clientes TABLE (Id INT PRIMARY KEY);
DECLARE @Reservas TABLE (Id INT PRIMARY KEY);
DECLARE @Comprobantes TABLE (Id INT PRIMARY KEY);
DECLARE @TiposDeporte TABLE (Id INT PRIMARY KEY);
DECLARE @TiposSuelo TABLE (Id INT PRIMARY KEY);
DECLARE @Desafios TABLE (Id INT PRIMARY KEY);

INSERT INTO @SuperAdmins (UsuarioId)
SELECT DISTINCT ur.UserId
FROM dbo.AspNetUserRoles ur
INNER JOIN dbo.AspNetRoles r ON r.Id = ur.RoleId
WHERE r.NormalizedName = N'OWNERPLATAFORMA'
   OR r.Name = N'OwnerPlataforma';

IF NOT EXISTS (SELECT 1 FROM @SuperAdmins)
BEGIN
    RAISERROR('No se encontro ningun usuario con rol OwnerPlataforma. Abortado para evitar eliminar todos los usuarios.', 16, 1);
    RETURN;
END;

INSERT INTO @UsuariosEliminar (UsuarioId)
SELECT u.Id
FROM dbo.AspNetUsers u
WHERE NOT EXISTS (
    SELECT 1
    FROM @SuperAdmins sa
    WHERE sa.UsuarioId = u.Id
);

INSERT INTO @Negocios (Id)
SELECT n.Id
FROM dbo.Negocios n;

INSERT INTO @Sedes (Id)
SELECT s.Id
FROM dbo.Sedes s
INNER JOIN @Negocios n ON n.Id = s.NegocioId;

INSERT INTO @Espacios (Id)
SELECT e.Id
FROM dbo.EspaciosDeportivos e
INNER JOIN @Sedes s ON s.Id = e.SedeId;

INSERT INTO @Clientes (Id)
SELECT c.Id
FROM dbo.Clientes c
INNER JOIN @Negocios n ON n.Id = c.NegocioId;

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
LEFT JOIN @Negocios n ON n.Id = ce.NegocioId
LEFT JOIN @Reservas r ON r.Id = ce.ReservaId
LEFT JOIN @Clientes c ON c.Id = ce.ClienteId
WHERE n.Id IS NOT NULL
   OR r.Id IS NOT NULL
   OR c.Id IS NOT NULL;

INSERT INTO @TiposDeporte (Id)
SELECT td.Id
FROM dbo.TiposDeporte td
INNER JOIN @Negocios n ON n.Id = td.NegocioId;

INSERT INTO @TiposSuelo (Id)
SELECT ts.Id
FROM dbo.TiposSuelo ts
INNER JOIN @Negocios n ON n.Id = ts.NegocioId;

INSERT INTO @Desafios (Id)
SELECT DISTINCT d.Id
FROM dbo.Desafio d
LEFT JOIN @UsuariosEliminar ue1 ON ue1.UsuarioId = d.IdUsuarioRetador
LEFT JOIN @UsuariosEliminar ue2 ON ue2.UsuarioId = d.IdUsuarioRetado
LEFT JOIN @TiposDeporte td ON td.Id = d.IdDeporte
WHERE ue1.UsuarioId IS NOT NULL
   OR ue2.UsuarioId IS NOT NULL
   OR td.Id IS NOT NULL;

DECLARE @Resumen TABLE
(
    Tabla NVARCHAR(128) NOT NULL,
    Filas INT NOT NULL
);

INSERT INTO @Resumen (Tabla, Filas) SELECT N'Negocios', COUNT(1) FROM @Negocios;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'Sedes', COUNT(1) FROM @Sedes;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'EspaciosDeportivos', COUNT(1) FROM @Espacios;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'Clientes', COUNT(1) FROM @Clientes;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'Reservas', COUNT(1) FROM @Reservas;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'ComprobantesElectronicos', COUNT(1) FROM @Comprobantes;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'TiposDeporte', COUNT(1) FROM @TiposDeporte;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'TiposSuelo', COUNT(1) FROM @TiposSuelo;
INSERT INTO @Resumen (Tabla, Filas) SELECT N'Desafio', COUNT(1) FROM @Desafios;

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'DesafioMensaje', COUNT(1)
FROM dbo.DesafioMensaje dm
WHERE EXISTS (SELECT 1 FROM @Desafios d WHERE d.Id = dm.IdDesafio)
   OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = dm.UsuarioIdEmisor);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'ReservasUsuariosPublicos', COUNT(1)
FROM dbo.ReservasUsuariosPublicos rup
WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = rup.ReservaId)
   OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = rup.UsuarioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SolicitudesReservaPublica', COUNT(1)
FROM dbo.SolicitudesReservaPublica srp
WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = srp.EspacioDeportivoId)
   OR EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = srp.ReservaId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'Pagos', COUNT(1)
FROM dbo.Pagos p
WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = p.ReservaId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'ComprobantesDetalle', COUNT(1)
FROM dbo.ComprobantesDetalle cd
WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = cd.ComprobanteElectronicoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'BloqueosHorario', COUNT(1)
FROM dbo.BloqueosHorario bh
WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = bh.EspacioDeportivoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'PromocionesHorario', COUNT(1)
FROM dbo.PromocionesHorario ph
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ph.NegocioId)
   OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ph.SedeId)
   OR EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = ph.EspacioDeportivoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'Tarifas', COUNT(1)
FROM dbo.Tarifas t
WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = t.EspacioDeportivoId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SedeConfiguracionNotificacion', COUNT(1)
FROM dbo.SedeConfiguracionNotificacion scn
WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = scn.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SedeFechasInhabilitadas', COUNT(1)
FROM dbo.SedeFechasInhabilitadas sfi
WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sfi.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SedeHorarioAtencion', COUNT(1)
FROM dbo.SedeHorarioAtencion sha
WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sha.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SedeServicios', COUNT(1)
FROM dbo.SedeServicios ss
WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ss.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SedesSeriesDocumentoComprobante', COUNT(1)
FROM dbo.SedesSeriesDocumentoComprobante ssd
WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ssd.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'SolicitudesAltaClub', COUNT(1)
FROM dbo.SolicitudesAltaClub sac
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = sac.NegocioId)
   OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sac.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'NegocioNotificaciones', COUNT(1)
FROM dbo.NegocioNotificaciones nn
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = nn.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'FormasPago', COUNT(1)
FROM dbo.FormasPago fp
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = fp.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'Monedas', COUNT(1)
FROM dbo.Monedas m
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = m.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'NegociosTiposDocumentoComprobante', COUNT(1)
FROM dbo.NegociosTiposDocumentoComprobante nd
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = nd.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'NegociosSeriesDocumentoComprobante', COUNT(1)
FROM dbo.NegociosSeriesDocumentoComprobante ns
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ns.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'NegociosSuscripcion', COUNT(1)
FROM dbo.NegociosSuscripcion ns
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ns.NegocioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'BitacoraAuditoria', COUNT(1)
FROM dbo.BitacoraAuditoria ba
WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ba.NegocioId)
   OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = ba.UsuarioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'UsuariosPublicosPerfil', COUNT(1)
FROM dbo.UsuariosPublicosPerfil upp
WHERE EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = upp.UsuarioId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'UsuariosNegocio', COUNT(1)
FROM dbo.UsuariosNegocio un
WHERE EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = un.UsuarioId)
   OR EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = un.NegocioId)
   OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = un.SedeId);

INSERT INTO @Resumen (Tabla, Filas)
SELECT N'AspNetUsers (a eliminar)', COUNT(1)
FROM @UsuariosEliminar;

PRINT '=== PREVIEW LIMPIEZA TOTAL ===';
SELECT COUNT(1) AS SuperAdminsConservados FROM @SuperAdmins;

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

    DELETE dm
    FROM dbo.DesafioMensaje dm
    WHERE EXISTS (SELECT 1 FROM @Desafios d WHERE d.Id = dm.IdDesafio)
       OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = dm.UsuarioIdEmisor);

    DELETE d
    FROM dbo.Desafio d
    WHERE EXISTS (SELECT 1 FROM @Desafios dx WHERE dx.Id = d.Id);

    DELETE rup
    FROM dbo.ReservasUsuariosPublicos rup
    WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = rup.ReservaId)
       OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = rup.UsuarioId);

    DELETE srp
    FROM dbo.SolicitudesReservaPublica srp
    WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = srp.EspacioDeportivoId)
       OR EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = srp.ReservaId);

    DELETE cd
    FROM dbo.ComprobantesDetalle cd
    WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = cd.ComprobanteElectronicoId);

    DELETE ce
    FROM dbo.ComprobantesElectronicos ce
    WHERE EXISTS (SELECT 1 FROM @Comprobantes c WHERE c.Id = ce.Id);

    DELETE p
    FROM dbo.Pagos p
    WHERE EXISTS (SELECT 1 FROM @Reservas r WHERE r.Id = p.ReservaId);

    DELETE bh
    FROM dbo.BloqueosHorario bh
    WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = bh.EspacioDeportivoId);

    DELETE ph
    FROM dbo.PromocionesHorario ph
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ph.NegocioId)
       OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ph.SedeId)
       OR EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = ph.EspacioDeportivoId);

    DELETE t
    FROM dbo.Tarifas t
    WHERE EXISTS (SELECT 1 FROM @Espacios e WHERE e.Id = t.EspacioDeportivoId);

    DELETE scn
    FROM dbo.SedeConfiguracionNotificacion scn
    WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = scn.SedeId);

    DELETE sfi
    FROM dbo.SedeFechasInhabilitadas sfi
    WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sfi.SedeId);

    DELETE sha
    FROM dbo.SedeHorarioAtencion sha
    WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sha.SedeId);

    DELETE ss
    FROM dbo.SedeServicios ss
    WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ss.SedeId);

    DELETE ssd
    FROM dbo.SedesSeriesDocumentoComprobante ssd
    WHERE EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = ssd.SedeId);

    DELETE sac
    FROM dbo.SolicitudesAltaClub sac
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = sac.NegocioId)
       OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = sac.SedeId);

    DELETE nn
    FROM dbo.NegocioNotificaciones nn
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = nn.NegocioId);

    DELETE fp
    FROM dbo.FormasPago fp
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = fp.NegocioId);

    DELETE m
    FROM dbo.Monedas m
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = m.NegocioId);

    DELETE nd
    FROM dbo.NegociosTiposDocumentoComprobante nd
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = nd.NegocioId);

    DELETE r
    FROM dbo.Reservas r
    WHERE EXISTS (SELECT 1 FROM @Reservas rx WHERE rx.Id = r.Id);

    DELETE c
    FROM dbo.Clientes c
    WHERE EXISTS (SELECT 1 FROM @Clientes cx WHERE cx.Id = c.Id);

    DELETE e
    FROM dbo.EspaciosDeportivos e
    WHERE EXISTS (SELECT 1 FROM @Espacios ex WHERE ex.Id = e.Id);

    DELETE s
    FROM dbo.Sedes s
    WHERE EXISTS (SELECT 1 FROM @Sedes sx WHERE sx.Id = s.Id);

    DELETE ba
    FROM dbo.BitacoraAuditoria ba
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ba.NegocioId)
       OR EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = ba.UsuarioId);

    DELETE upp
    FROM dbo.UsuariosPublicosPerfil upp
    WHERE EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = upp.UsuarioId);

    DELETE un
    FROM dbo.UsuariosNegocio un
    WHERE EXISTS (SELECT 1 FROM @UsuariosEliminar u WHERE u.UsuarioId = un.UsuarioId)
       OR EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = un.NegocioId)
       OR EXISTS (SELECT 1 FROM @Sedes s WHERE s.Id = un.SedeId);

    DELETE ns
    FROM dbo.NegociosSeriesDocumentoComprobante ns
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = ns.NegocioId);

    DELETE nsub
    FROM dbo.NegociosSuscripcion nsub
    WHERE EXISTS (SELECT 1 FROM @Negocios n WHERE n.Id = nsub.NegocioId);

    DELETE td
    FROM dbo.TiposDeporte td
    WHERE EXISTS (SELECT 1 FROM @TiposDeporte tx WHERE tx.Id = td.Id);

    DELETE ts
    FROM dbo.TiposSuelo ts
    WHERE EXISTS (SELECT 1 FROM @TiposSuelo tx WHERE tx.Id = ts.Id);

    DELETE n
    FROM dbo.Negocios n
    WHERE EXISTS (SELECT 1 FROM @Negocios nx WHERE nx.Id = n.Id);

    DELETE u
    FROM dbo.AspNetUsers u
    WHERE EXISTS (SELECT 1 FROM @UsuariosEliminar ux WHERE ux.UsuarioId = u.Id);

    COMMIT;
    PRINT 'Limpieza total ejecutada correctamente. Solo se conservaron usuarios superadmin.';
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

