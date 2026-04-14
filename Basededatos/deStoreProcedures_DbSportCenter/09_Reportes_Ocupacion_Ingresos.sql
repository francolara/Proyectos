-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Reportes de ocupacion e ingresos por rango y habilitacion del modulo REPORTES.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   13/04/2026
-- Description:   Reportes v2: filtro por sede, ocupacion con IDs para drill-down y resumen operativo; excluye canceladas (Estado=5) en KPIs de cantidad/montos.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_SeedModulosPermisosBase
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'DASHBOARD')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'DASHBOARD', N'Dashboard', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'SEDES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'SEDES', N'Sedes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'CLIENTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'CLIENTES', N'Clientes', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'ESPACIOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'ESPACIOS', N'Espacios deportivos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'RESERVAS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'RESERVAS', N'Reservas', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'PAGOS')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'PAGOS', N'Pagos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'COMPROBANTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'COMPROBANTES', N'Comprobantes electronicos', 1);
        IF NOT EXISTS (SELECT 1 FROM dbo.ModulosSistema WHERE Codigo = N'REPORTES')
            INSERT INTO dbo.ModulosSistema (Codigo, Nombre, Activo) VALUES (N'REPORTES', N'Reportes', 1);

        ;WITH Roles AS
        (
            SELECT CAST(1 AS INT) AS RolNegocio UNION ALL
            SELECT 2 UNION ALL
            SELECT 3 UNION ALL
            SELECT 4 UNION ALL
            SELECT 5
        )
        INSERT INTO dbo.RolesNegocioPermiso (RolNegocio, ModuloSistemaId, PuedeVer, PuedeCrear, PuedeEditar, PuedeEliminar)
        SELECT
            r.RolNegocio,
            m.Id,
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'PAGOS', N'CLIENTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'DASHBOARD', N'PAGOS', N'COMPROBANTES', N'REPORTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'DASHBOARD', N'RESERVAS', N'ESPACIOS', N'REPORTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1
                      WHEN r.RolNegocio = 2 AND m.Codigo IN (N'RESERVAS', N'PAGOS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 3 AND m.Codigo IN (N'RESERVAS', N'CLIENTES') THEN 1
                      WHEN r.RolNegocio = 4 AND m.Codigo IN (N'PAGOS', N'COMPROBANTES') THEN 1
                      WHEN r.RolNegocio = 5 AND m.Codigo IN (N'RESERVAS', N'ESPACIOS') THEN 1
                      ELSE 0 END AS BIT),
            CAST(CASE WHEN r.RolNegocio = 1 THEN 1 ELSE 0 END AS BIT)
        FROM Roles r
        CROSS JOIN dbo.ModulosSistema m
        WHERE NOT EXISTS (
            SELECT 1
            FROM dbo.RolesNegocioPermiso rp
            WHERE rp.RolNegocio = r.RolNegocio
              AND rp.ModuloSistemaId = m.Id
        );
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_OcupacionPorEspacio
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH PagosPorReserva AS
        (
            SELECT
                p.ReservaId,
                SUM(p.Monto) AS MontoCobrado
            FROM dbo.Pagos p
            GROUP BY p.ReservaId
        )
        SELECT
            s.Id AS SedeId,
            e.Id AS EspacioDeportivoId,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            COUNT(1) AS CantidadReservas,
            CAST(SUM(DATEDIFF(MINUTE, r.HoraInicio, r.HoraFin)) / 60.0 AS DECIMAL(10,2)) AS HorasReservadas,
            SUM(r.Total) AS MontoReservado,
            SUM(COALESCE(pr.MontoCobrado, 0)) AS MontoCobrado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN PagosPorReserva pr ON pr.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado NOT IN (5, 6)
        GROUP BY s.Id, e.Id, s.Nombre, e.Nombre
        ORDER BY HorasReservadas DESC, CantidadReservas DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_IngresosPorDia
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Fecha,
            COUNT(DISTINCT r.Id) AS CantidadReservas,
            COALESCE(SUM(p.Monto), 0) AS Ingresos
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado <> 5
        GROUP BY r.Fecha
        ORDER BY r.Fecha ASC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reportes_ResumenOperativo
    @NegocioId INT,
    @FechaDesde DATE,
    @FechaHasta DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        ;WITH ReservasBase AS
        (
            SELECT
                r.Id,
                r.Estado,
                r.Total
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND r.Fecha >= @FechaDesde
              AND r.Fecha <= @FechaHasta
        ),
        PagosPorReserva AS
        (
            SELECT
                p.ReservaId,
                SUM(p.Monto) AS MontoCobrado
            FROM dbo.Pagos p
            INNER JOIN ReservasBase rb ON rb.Id = p.ReservaId
            GROUP BY p.ReservaId
        )
        SELECT
            COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN 1 ELSE 0 END), 0) AS TotalReservas,
            COALESCE(SUM(CASE WHEN rb.Estado = 1 THEN 1 ELSE 0 END), 0) AS TotalPendientes,
            COALESCE(SUM(CASE WHEN rb.Estado IN (2, 3) THEN 1 ELSE 0 END), 0) AS TotalConfirmadas,
            COALESCE(SUM(CASE WHEN rb.Estado = 4 THEN 1 ELSE 0 END), 0) AS TotalPagadas,
            COALESCE(SUM(CASE WHEN rb.Estado = 5 THEN 1 ELSE 0 END), 0) AS TotalCanceladas,
            COALESCE(SUM(CASE WHEN rb.Estado = 6 THEN 1 ELSE 0 END), 0) AS TotalNoShow,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN rb.Total ELSE 0 END), 0) AS DECIMAL(18,2)) AS MontoReservado,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 THEN COALESCE(pr.MontoCobrado, 0) ELSE 0 END), 0) AS DECIMAL(18,2)) AS MontoCobrado,
            CAST(COALESCE(SUM(CASE WHEN rb.Estado <> 5 AND rb.Total - COALESCE(pr.MontoCobrado, 0) > 0 THEN rb.Total - COALESCE(pr.MontoCobrado, 0) ELSE 0 END), 0) AS DECIMAL(18,2)) AS SaldoPendiente
        FROM ReservasBase rb
        LEFT JOIN PagosPorReserva pr ON pr.ReservaId = rb.Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
