-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/03/2026
-- Description:   Asignacion de sede por usuario no administrador y restriccion de filtros por sede en backend (incluye listado de usuarios filtrado por sede).
-- Firma:         Codex - 29/03/2026
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/03/2026
-- Description:   Sp_Espacios_Listar ahora devuelve TarifaResumen calculado en backend (SQL) por dias, horarios y moneda del negocio; Sp_UsuariosNegocio_ActualizarRol valida filas afectadas; Sp_Combos_EspaciosPorNegocio muestra Codigo + Nombre + (Tipo suelo).
-- Firma:         Codex - 30/03/2026 | TarifaResumen backend + validacion estricta de actualizacion de rol por negocio + formato de combo de espacios para reservas.
-- =============================================

IF COL_LENGTH('dbo.UsuariosNegocio', 'SedeId') IS NULL
BEGIN
    ALTER TABLE dbo.UsuariosNegocio ADD SedeId INT NULL;
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_UsuariosNegocio_Sedes_SedeId')
BEGIN
    ALTER TABLE dbo.UsuariosNegocio
    ADD CONSTRAINT FK_UsuariosNegocio_Sedes_SedeId
        FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id);
END;
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.UsuariosNegocio') AND name = N'IX_UsuariosNegocio_SedeId')
BEGIN
    CREATE INDEX IX_UsuariosNegocio_SedeId ON dbo.UsuariosNegocio (SedeId);
END;
GO

;WITH SedeDefault AS
(
    SELECT s.NegocioId, MIN(s.Id) AS SedeId
    FROM dbo.Sedes s
    GROUP BY s.NegocioId
)
UPDATE un
SET SedeId = sd.SedeId
FROM dbo.UsuariosNegocio un
INNER JOIN SedeDefault sd ON sd.NegocioId = un.NegocioId
WHERE un.RolNegocio <> 1
  AND un.SedeId IS NULL;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_Sedes
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT s.Id, s.Nombre
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND s.Activo = 1
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id,
            s.Nombre,
            s.Direccion,
            STUFF((
                SELECT N', ' + cs.Nombre
                FROM dbo.SedeServicios ss
                INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = ss.ServicioId
                WHERE ss.SedeId = s.Id
                  AND cs.Activo = 1
                ORDER BY cs.Nombre
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 2, N'') AS Servicios,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            CONCAT(
                CASE WHEN COALESCE(sha.AtiendeLunes, 1) = 1 THEN N'Lun ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeMartes, 1) = 1 THEN N'Mar ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeMiercoles, 1) = 1 THEN N'Mie ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeJueves, 1) = 1 THEN N'Jue ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeViernes, 1) = 1 THEN N'Vie ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeSabado, 1) = 1 THEN N'Sab ' ELSE N'' END,
                CASE WHEN COALESCE(sha.AtiendeDomingo, 1) = 1 THEN N'Dom' ELSE N'' END
            ) AS DiasAtencion,
            CONCAT(CONVERT(NVARCHAR(5), COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)), 108), N' - ', CONVERT(NVARCHAR(5), COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)), 108)) AS HorarioAtencion,
            (SELECT COUNT(1) FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = s.Id AND sfi.Activo = 1) AS FechasNoLaborablesCount,
            s.Activo
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Espacios_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @SimboloMoneda NVARCHAR(10);
        SET @SimboloMoneda = N'S/';

        SELECT TOP (1) @SimboloMoneda = COALESCE(m.Simbolo, N'S/')
        FROM dbo.Negocios n
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        WHERE n.Id = @NegocioId;

        SELECT
            e.Id,
            e.Codigo,
            e.Nombre,
            s.Nombre AS Sede,
            td.Nombre AS TipoDeporte,
            ts.Nombre AS TipoSuelo,
            CASE e.Estado WHEN 1 THEN N'Activo' WHEN 2 THEN N'EnMantenimiento' ELSE N'Inactivo' END AS Estado,
            COALESCE
            (
                NULLIF
                (
                    STUFF
                    (
                        (
                            SELECT N' | '
                                + CASE t.DiaSemana
                                    WHEN 1 THEN N'Lun'
                                    WHEN 2 THEN N'Mar'
                                    WHEN 3 THEN N'Mie'
                                    WHEN 4 THEN N'Jue'
                                    WHEN 5 THEN N'Vie'
                                    WHEN 6 THEN N'Sab'
                                    WHEN 0 THEN N'Dom'
                                    ELSE N'Dia'
                                  END
                                + N' '
                                + CONVERT(NVARCHAR(5), t.HoraInicio, 108)
                                + N'-'
                                + CONVERT(NVARCHAR(5), t.HoraFin, 108)
                                + N' '
                                + @SimboloMoneda
                                + CONVERT(NVARCHAR(20), CAST(t.Precio AS DECIMAL(10,2)))
                            FROM dbo.Tarifas t
                            WHERE t.EspacioDeportivoId = e.Id
                              AND t.Activa = 1
                            ORDER BY
                                CASE t.DiaSemana
                                    WHEN 1 THEN 1
                                    WHEN 2 THEN 2
                                    WHEN 3 THEN 3
                                    WHEN 4 THEN 4
                                    WHEN 5 THEN 5
                                    WHEN 6 THEN 6
                                    WHEN 0 THEN 7
                                    ELSE 8
                                END,
                                t.HoraInicio,
                                t.HoraFin
                            FOR XML PATH(''), TYPE
                        ).value('.', 'NVARCHAR(MAX)'),
                        1, 3, N''
                    ),
                    N''
                ),
                N'Sin tarifa configurada'
            ) AS TarifaResumen
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_EspaciosPorNegocio
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            CONCAT(
                COALESCE(NULLIF(LTRIM(RTRIM(e.Codigo)), N''), N'S/C'),
                N' - ',
                e.Nombre,
                N' (',
                COALESCE(ts.Nombre, N'Sin suelo'),
                N')'
            )
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND e.Estado = 1
        ORDER BY e.Codigo, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Combos_ReservasPorNegocio
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT r.Id, CONCAT(N'#', r.Id, N' - ', c.NombresORazonSocial)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY r.Fecha DESC, r.HoraInicio DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Pagos_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            p.Id,
            p.ReservaId,
            p.FechaPago,
            p.Monto,
            CAST(p.FormaPago AS NVARCHAR(20))
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY p.FechaPago DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Comprobantes_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT TOP (100)
            c.Id,
            CAST(c.TipoComprobante AS NVARCHAR(20)),
            CONCAT(c.Serie, N'-', c.Numero),
            c.FechaEmision,
            cl.NombresORazonSocial,
            c.Total,
            CAST(c.Estado AS NVARCHAR(20))
        FROM dbo.ComprobantesElectronicos c
        INNER JOIN dbo.Clientes cl ON cl.Id = c.ClienteId
        LEFT JOIN dbo.Reservas r ON r.Id = c.ReservaId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        LEFT JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE c.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY c.FechaEmision DESC;
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
        SELECT
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            COUNT(1) AS CantidadReservas,
            CAST(SUM(DATEDIFF(MINUTE, r.HoraInicio, r.HoraFin)) / 60.0 AS DECIMAL(10,2)) AS HorasReservadas,
            SUM(r.Total) AS MontoReservado,
            COALESCE(SUM(p.Monto), 0) AS MontoCobrado
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.Pagos p ON p.ReservaId = r.Id
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha >= @FechaDesde
          AND r.Fecha <= @FechaHasta
          AND r.Estado NOT IN (5, 6)
        GROUP BY s.Nombre, e.Nombre
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

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ObtenerMetricas
    @NegocioId INT,
    @Fecha DATE,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @TotalSedes INT = 0, @TotalEspacios INT = 0, @ReservasHoy INT = 0;
        DECLARE @IngresosHoy DECIMAL(12,2) = 0, @OcupacionHoyPct DECIMAL(5,2) = 0;
        DECLARE @NoShowMes INT = 0, @TicketPromedioMes DECIMAL(12,2) = 0;

        SELECT @TotalSedes = COUNT(1)
        FROM dbo.Sedes s
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId);

        SELECT @TotalEspacios = COUNT(1)
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId);

        SELECT @ReservasHoy = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Fecha = @Fecha;

        SELECT @IngresosHoy = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND CAST(p.FechaPago AS DATE) = @Fecha;

        IF @TotalEspacios > 0
        BEGIN
            DECLARE @EspaciosOcupadosHoy INT = 0;
            SELECT @EspaciosOcupadosHoy = COUNT(DISTINCT r.EspacioDeportivoId)
            FROM dbo.Reservas r
            INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
            INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
            WHERE s.NegocioId = @NegocioId
              AND (@SedeId IS NULL OR s.Id = @SedeId)
              AND r.Fecha = @Fecha
              AND r.Estado NOT IN (5, 6);
            SET @OcupacionHoyPct = CAST((@EspaciosOcupadosHoy * 100.0) / @TotalEspacios AS DECIMAL(5,2));
        END;

        SELECT @NoShowMes = COUNT(1)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND r.Estado = 6
          AND YEAR(r.Fecha) = YEAR(@Fecha)
          AND MONTH(r.Fecha) = MONTH(@Fecha);

        DECLARE @TotalCobradoMes DECIMAL(12,2) = 0, @ReservasPagadasMes INT = 0;

        SELECT @TotalCobradoMes = COALESCE(SUM(p.Monto), 0)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        SELECT @ReservasPagadasMes = COUNT(DISTINCT p.ReservaId)
        FROM dbo.Pagos p
        INNER JOIN dbo.Reservas r ON r.Id = p.ReservaId
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
          AND YEAR(p.FechaPago) = YEAR(@Fecha)
          AND MONTH(p.FechaPago) = MONTH(@Fecha);

        IF @ReservasPagadasMes > 0
            SET @TicketPromedioMes = CAST(@TotalCobradoMes / @ReservasPagadasMes AS DECIMAL(12,2));

        SELECT
            @TotalSedes AS TotalSedes,
            @TotalEspacios AS TotalEspacios,
            @ReservasHoy AS ReservasHoy,
            @IngresosHoy AS IngresosHoy,
            @OcupacionHoyPct AS OcupacionHoyPct,
            @NoShowMes AS NoShowMes,
            @TicketPromedioMes AS TicketPromedioMes;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Promociones_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            p.Id,
            p.Nombre,
            COALESCE(s.Nombre, N'Todas') AS Sede,
            COALESCE(e.Nombre, N'Todos') AS Espacio,
            p.FechaInicio,
            p.FechaFin,
            p.HoraInicio,
            p.HoraFin,
            p.PorcentajeDescuento,
            p.Activo
        FROM dbo.PromocionesHorario p
        LEFT JOIN dbo.Sedes s ON s.Id = p.SedeId
        LEFT JOIN dbo.EspaciosDeportivos e ON e.Id = p.EspacioDeportivoId
        WHERE p.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR p.SedeId = @SedeId OR (p.SedeId IS NULL AND p.EspacioDeportivoId IS NULL))
        ORDER BY p.FechaInicio DESC, p.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_Listar
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            un.Id AS UsuarioNegocioId,
            un.UsuarioId,
            COALESCE(u.Nombres, N'') AS Nombres,
            COALESCE(u.Apellidos, N'') AS Apellidos,
            COALESCE(u.Email, N'') AS Correo,
            un.RolNegocio,
            un.Activo,
            un.SedeId,
            COALESCE(s.Nombre, N'') AS SedeNombre
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.AspNetUsers u ON u.Id = un.UsuarioId
        LEFT JOIN dbo.Sedes s ON s.Id = un.SedeId
        WHERE un.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR un.SedeId = @SedeId)
        ORDER BY un.Activo DESC, u.Email;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_AsignarPorCorreo
    @NegocioId INT,
    @Correo NVARCHAR(256),
    @RolNegocio INT,
    @SedeId INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioId NVARCHAR(450);
        SELECT TOP (1) @UsuarioId = u.Id
        FROM dbo.AspNetUsers u
        WHERE u.NormalizedEmail = UPPER(@Correo);

        IF @UsuarioId IS NULL
            RAISERROR('No existe usuario con ese correo en el sistema.', 16, 1);

        IF @RolNegocio = 1
            SET @SedeId = NULL;

        IF @RolNegocio <> 1 AND @SedeId IS NULL
            RAISERROR('La sede es obligatoria para usuarios no administradores.', 16, 1);

        IF @SedeId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.Sedes s WHERE s.Id = @SedeId AND s.NegocioId = @NegocioId)
                RAISERROR('La sede no pertenece al negocio seleccionado.', 16, 1);
        END;

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId)
        BEGIN
            UPDATE dbo.UsuariosNegocio
            SET RolNegocio = @RolNegocio,
                SedeId = @SedeId,
                Activo = 1
            WHERE NegocioId = @NegocioId
              AND UsuarioId = @UsuarioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, SedeId, Activo)
            VALUES (@UsuarioId, @NegocioId, @RolNegocio, @SedeId, 1);
        END;

        DECLARE @UsuarioNegocioId INT;
        SELECT TOP (1) @UsuarioNegocioId = Id FROM dbo.UsuariosNegocio WHERE NegocioId = @NegocioId AND UsuarioId = @UsuarioId;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'CREATE', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_UsuariosNegocio_ActualizarRol
    @NegocioId INT,
    @UsuarioNegocioId INT,
    @RolNegocio INT,
    @SedeId INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @RolNegocio = 1
            SET @SedeId = NULL;

        IF @RolNegocio <> 1 AND @SedeId IS NULL
            RAISERROR('La sede es obligatoria para usuarios no administradores.', 16, 1);

        IF @SedeId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.Sedes s WHERE s.Id = @SedeId AND s.NegocioId = @NegocioId)
                RAISERROR('La sede no pertenece al negocio seleccionado.', 16, 1);
        END;

        UPDATE dbo.UsuariosNegocio
        SET RolNegocio = @RolNegocio,
            SedeId = @SedeId
        WHERE Id = @UsuarioNegocioId
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el usuario del negocio para actualizar rol.', 16, 1);

        IF @@ROWCOUNT > 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @UsuarioNegocioId);
            EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'USUARIOS', @Accion = N'EDIT', @Entidad = N'UsuarioNegocio', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Seguridad_ObtenerContextoModulo
    @UsuarioId NVARCHAR(450),
    @NegocioId INT,
    @ModuloCodigo NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT, @NegocioNombre NVARCHAR(200), @ModuloId INT, @ModuloNombre NVARCHAR(120);
        DECLARE @PuedeVer BIT = 0, @PuedeCrear BIT = 0, @PuedeEditar BIT = 0, @PuedeEliminar BIT = 0;
        DECLARE @EstadoSuscripcion INT, @EsPrueba BIT, @FechaFinPrueba DATE, @FechaFinPlan DATE;
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);
        DECLARE @SedeIdAsignada INT = NULL, @EsAdministrador BIT = 0;

        SELECT
            @UsuarioNegocioId = un.Id,
            @RolNegocio = un.RolNegocio,
            @SedeIdAsignada = un.SedeId,
            @NegocioNombre = n.NombreComercial
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1
          AND n.Activo = 1;

        SET @EsAdministrador = CASE WHEN @RolNegocio = 1 THEN 1 ELSE 0 END;
        IF @EsAdministrador = 1
            SET @SedeIdAsignada = NULL;

        IF @UsuarioNegocioId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, N'', @ModuloCodigo, N'', N'', CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Usuario sin acceso al negocio', CAST(NULL AS INT), CAST(0 AS BIT);
            RETURN;
        END;

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NOT NULL
        BEGIN
            SELECT
                @EstadoSuscripcion = ns.EstadoSuscripcion,
                @EsPrueba = ns.EsPrueba,
                @FechaFinPrueba = ns.FechaFinPrueba,
                @FechaFinPlan = ns.FechaFinPlan
            FROM dbo.NegociosSuscripcion ns
            WHERE ns.NegocioId = @NegocioId;

            IF @EstadoSuscripcion = 1 AND @EsPrueba = 1 AND @FechaFinPrueba IS NOT NULL AND @FechaFinPrueba < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    EsPrueba = 0,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion = 2 AND @FechaFinPlan IS NOT NULL AND @FechaFinPlan < @Hoy
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 3,
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @UsuarioId
                WHERE NegocioId = @NegocioId;
                SET @EstadoSuscripcion = 3;
            END;

            IF @EstadoSuscripcion IN (3, 4)
            BEGIN
                SELECT
                    CAST(0 AS BIT),
                    @NegocioId,
                    @NegocioNombre,
                    @ModuloCodigo,
                    N'',
                    CAST(@RolNegocio AS NVARCHAR(20)),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    CAST(0 AS BIT),
                    N'La suscripcion del negocio esta vencida o suspendida. Activa un plan para continuar operando.',
                    @SedeIdAsignada,
                    @EsAdministrador;
                RETURN;
            END;
        END;

        SELECT @ModuloId = m.Id, @ModuloNombre = m.Nombre
        FROM dbo.ModulosSistema m
        WHERE m.Codigo = @ModuloCodigo AND m.Activo = 1;

        IF @ModuloId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, @NegocioNombre, @ModuloCodigo, N'', CAST(@RolNegocio AS NVARCHAR(20)), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Modulo no configurado', @SedeIdAsignada, @EsAdministrador;
            RETURN;
        END;

        SELECT
            @PuedeVer = rp.PuedeVer,
            @PuedeCrear = rp.PuedeCrear,
            @PuedeEditar = rp.PuedeEditar,
            @PuedeEliminar = rp.PuedeEliminar
        FROM dbo.RolesNegocioPermiso rp
        WHERE rp.RolNegocio = @RolNegocio
          AND rp.ModuloSistemaId = @ModuloId;

        SELECT
            @PuedeVer = COALESCE(up.PuedeVer, @PuedeVer),
            @PuedeCrear = COALESCE(up.PuedeCrear, @PuedeCrear),
            @PuedeEditar = COALESCE(up.PuedeEditar, @PuedeEditar),
            @PuedeEliminar = COALESCE(up.PuedeEliminar, @PuedeEliminar)
        FROM dbo.UsuariosNegocioPermiso up
        WHERE up.UsuarioNegocioId = @UsuarioNegocioId
          AND up.ModuloSistemaId = @ModuloId;

        SELECT
            CAST(CASE WHEN @PuedeVer = 1 THEN 1 ELSE 0 END AS BIT) AS Autorizado,
            @NegocioId,
            @NegocioNombre,
            @ModuloCodigo,
            @ModuloNombre,
            CAST(@RolNegocio AS NVARCHAR(20)) AS RolActual,
            @PuedeVer,
            @PuedeCrear,
            @PuedeEditar,
            @PuedeEliminar,
            CAST(NULL AS NVARCHAR(200)) AS Mensaje,
            @SedeIdAsignada AS SedeIdAsignada,
            @EsAdministrador AS EsAdministrador;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
