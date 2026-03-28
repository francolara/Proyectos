-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Dias/horario de atencion por sede + fechas no laborables y validacion de reservas.
-- Firma:         Codex - 27/03/2026
-- =============================================

IF OBJECT_ID(N'dbo.SedeHorarioAtencion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SedeHorarioAtencion
    (
        SedeId INT NOT NULL PRIMARY KEY,
        AtiendeLunes BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Lunes DEFAULT (1),
        AtiendeMartes BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Martes DEFAULT (1),
        AtiendeMiercoles BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Miercoles DEFAULT (1),
        AtiendeJueves BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Jueves DEFAULT (1),
        AtiendeViernes BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Viernes DEFAULT (1),
        AtiendeSabado BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Sabado DEFAULT (1),
        AtiendeDomingo BIT NOT NULL CONSTRAINT DF_SedeHorarioAtencion_Domingo DEFAULT (1),
        HoraApertura TIME NOT NULL CONSTRAINT DF_SedeHorarioAtencion_HoraApertura DEFAULT ('08:00'),
        HoraCierre TIME NOT NULL CONSTRAINT DF_SedeHorarioAtencion_HoraCierre DEFAULT ('23:00'),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_SedeHorarioAtencion_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_SedeHorarioAtencion_Sede FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id)
    );
END;
GO

IF OBJECT_ID(N'dbo.SedeFechasInhabilitadas', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SedeFechasInhabilitadas
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        SedeId INT NOT NULL,
        Fecha DATE NOT NULL,
        Activo BIT NOT NULL CONSTRAINT DF_SedeFechasInhabilitadas_Activo DEFAULT (1),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_SedeFechasInhabilitadas_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_SedeFechasInhabilitadas_Sede FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id),
        CONSTRAINT UQ_SedeFechasInhabilitadas_Sede_Fecha UNIQUE (SedeId, Fecha)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Listar
    @NegocioId INT
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
        ORDER BY s.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_ObtenerPorId
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            s.Id, s.NegocioId, s.Nombre, s.Direccion, s.Telefono, s.Activo,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(20), ss.ServicioId) FROM dbo.SedeServicios ss WHERE ss.SedeId = s.Id ORDER BY ss.ServicioId FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS ServiciosIdsCsv,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp,
            COALESCE(sha.AtiendeLunes, 1) AS AtiendeLunes,
            COALESCE(sha.AtiendeMartes, 1) AS AtiendeMartes,
            COALESCE(sha.AtiendeMiercoles, 1) AS AtiendeMiercoles,
            COALESCE(sha.AtiendeJueves, 1) AS AtiendeJueves,
            COALESCE(sha.AtiendeViernes, 1) AS AtiendeViernes,
            COALESCE(sha.AtiendeSabado, 1) AS AtiendeSabado,
            COALESCE(sha.AtiendeDomingo, 1) AS AtiendeDomingo,
            COALESCE(sha.HoraApertura, CAST('08:00' AS TIME)) AS HoraApertura,
            COALESCE(sha.HoraCierre, CAST('23:00' AS TIME)) AS HoraCierre,
            STUFF((SELECT N',' + CONVERT(NVARCHAR(10), sfi.Fecha, 23) FROM dbo.SedeFechasInhabilitadas sfi WHERE sfi.SedeId = s.Id AND sfi.Activo = 1 ORDER BY sfi.Fecha FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS FechasInhabilitadasCsv
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        LEFT JOIN dbo.SedeHorarioAtencion sha ON sha.SedeId = s.Id
        WHERE s.NegocioId = @NegocioId
          AND s.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
