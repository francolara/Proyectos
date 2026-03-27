-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 7.1 - Configuracion de notificaciones y contacto por sede.
-- =============================================

IF OBJECT_ID(N'dbo.SedeConfiguracionNotificacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SedeConfiguracionNotificacion
    (
        SedeId INT NOT NULL PRIMARY KEY,
        NotificacionesActivas BIT NOT NULL CONSTRAINT DF_SedeConfigNotif_Activas DEFAULT (1),
        MinutosAnticipacionRecordatorio INT NOT NULL CONSTRAINT DF_SedeConfigNotif_Anticipacion DEFAULT (90),
        MinutosToleranciaNoShow INT NOT NULL CONSTRAINT DF_SedeConfigNotif_NoShow DEFAULT (30),
        CorreoNotificacion NVARCHAR(200) NULL,
        WhatsappContacto NVARCHAR(20) NULL,
        PermiteChatWhatsapp BIT NOT NULL CONSTRAINT DF_SedeConfigNotif_ChatWhatsapp DEFAULT (0),
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_SedeConfigNotif_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_SedeConfigNotif_Sedes_SedeId FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id)
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
            s.Activo
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
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
            s.Id,
            s.NegocioId,
            s.Nombre,
            s.Direccion,
            s.Telefono,
            s.Activo,
            STUFF((
                SELECT N',' + CONVERT(NVARCHAR(20), ss.ServicioId)
                FROM dbo.SedeServicios ss
                WHERE ss.SedeId = s.Id
                ORDER BY ss.ServicioId
                FOR XML PATH(''), TYPE
            ).value('.', 'NVARCHAR(MAX)'), 1, 1, N'') AS ServiciosIdsCsv,
            COALESCE(scn.NotificacionesActivas, 1) AS NotificacionesActivas,
            COALESCE(scn.MinutosAnticipacionRecordatorio, 90) AS MinutosAnticipacionRecordatorio,
            COALESCE(scn.MinutosToleranciaNoShow, 30) AS MinutosToleranciaNoShow,
            scn.CorreoNotificacion,
            scn.WhatsappContacto,
            COALESCE(scn.PermiteChatWhatsapp, 0) AS PermiteChatWhatsapp
        FROM dbo.Sedes s
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
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

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Crear
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @NotificacionesActivas BIT = 1,
    @MinutosAnticipacionRecordatorio INT = 90,
    @MinutosToleranciaNoShow INT = 30,
    @CorreoNotificacion NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(20) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @MinutosAnticipacionRecordatorio < 5 OR @MinutosAnticipacionRecordatorio > 1440
            RAISERROR('Minutos de anticipacion fuera de rango (5 a 1440).', 16, 1);
        IF @MinutosToleranciaNoShow < 0 OR @MinutosToleranciaNoShow > 240
            RAISERROR('Minutos de tolerancia no-show fuera de rango (0 a 240).', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, @Nombre, @Direccion, @Telefono, @Activo, SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        INSERT INTO dbo.SedeConfiguracionNotificacion
        (
            SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
            CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @Id, @NotificacionesActivas, @MinutosAnticipacionRecordatorio, @MinutosToleranciaNoShow,
            @CorreoNotificacion, @WhatsappContacto, @PermiteChatWhatsapp, SYSUTCDATETIME(), @Usuario
        );

        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'SEDES',
            @Accion = N'CREATE',
            @Entidad = N'Sede',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_Actualizar
    @Id INT,
    @NegocioId INT,
    @Nombre NVARCHAR(150),
    @Direccion NVARCHAR(250),
    @Telefono NVARCHAR(20) = NULL,
    @Activo BIT,
    @ServiciosIdsCsv NVARCHAR(MAX) = NULL,
    @NotificacionesActivas BIT = 1,
    @MinutosAnticipacionRecordatorio INT = 90,
    @MinutosToleranciaNoShow INT = 30,
    @CorreoNotificacion NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(20) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @MinutosAnticipacionRecordatorio < 5 OR @MinutosAnticipacionRecordatorio > 1440
            RAISERROR('Minutos de anticipacion fuera de rango (5 a 1440).', 16, 1);
        IF @MinutosToleranciaNoShow < 0 OR @MinutosToleranciaNoShow > 240
            RAISERROR('Minutos de tolerancia no-show fuera de rango (0 a 240).', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Sedes
        SET Nombre = @Nombre,
            Direccion = @Direccion,
            Telefono = @Telefono,
            Activo = @Activo,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        IF @@ROWCOUNT = 0
        BEGIN
            ROLLBACK TRANSACTION;
            RETURN;
        END;

        MERGE dbo.SedeConfiguracionNotificacion AS tgt
        USING (SELECT @Id AS SedeId) AS src
            ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN
            UPDATE SET
                NotificacionesActivas = @NotificacionesActivas,
                MinutosAnticipacionRecordatorio = @MinutosAnticipacionRecordatorio,
                MinutosToleranciaNoShow = @MinutosToleranciaNoShow,
                CorreoNotificacion = @CorreoNotificacion,
                WhatsappContacto = @WhatsappContacto,
                PermiteChatWhatsapp = @PermiteChatWhatsapp,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT
            (
                SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
                CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, @NotificacionesActivas, @MinutosAnticipacionRecordatorio, @MinutosToleranciaNoShow,
                @CorreoNotificacion, @WhatsappContacto, @PermiteChatWhatsapp, SYSUTCDATETIME(), @Usuario
            );

        DELETE ss
        FROM dbo.SedeServicios ss
        WHERE ss.SedeId = @Id;

        IF @ServiciosIdsCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@ServiciosIdsCsv))) > 0
        BEGIN
            ;WITH Servicios AS
            (
                SELECT DISTINCT TRY_CONVERT(INT, LTRIM(RTRIM(value))) AS ServicioId
                FROM STRING_SPLIT(@ServiciosIdsCsv, N',')
                WHERE TRY_CONVERT(INT, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeServicios (SedeId, ServicioId, FechaRegistro, UsuarioCreacion)
            SELECT @Id, s.ServicioId, SYSUTCDATETIME(), @Usuario
            FROM Servicios s
            INNER JOIN dbo.CatalogoServiciosSede cs ON cs.Id = s.ServicioId
            WHERE cs.Activo = 1;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'SEDES',
            @Accion = N'EDIT',
            @Entidad = N'Sede',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_RecordatoriosPendientes
    @FechaHoraActual DATETIME2
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id AS ReservaId,
            s.NegocioId,
            c.NombresORazonSocial AS Cliente,
            c.Correo,
            s.Nombre AS Sede,
            e.Nombre AS Espacio,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            scn.CorreoNotificacion,
            scn.WhatsappContacto
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.Clientes c ON c.Id = r.ClienteId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE r.Estado IN (1, 2)
          AND r.RecordatorioEnviado = 0
          AND c.Correo IS NOT NULL
          AND LTRIM(RTRIM(c.Correo)) <> N''
          AND COALESCE(scn.NotificacionesActivas, 1) = 1
          AND @FechaHoraActual >= DATEADD(
                MINUTE,
                -COALESCE(scn.MinutosAnticipacionRecordatorio, 90),
                DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2))
          )
          AND @FechaHoraActual <= DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_AutoNoShow
    @FechaHoraActual DATETIME2,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Actualizadas TABLE
        (
            ReservaId INT NOT NULL,
            NegocioId INT NOT NULL
        );

        UPDATE r
        SET r.Estado = 6,
            r.FechaActualizacion = SYSUTCDATETIME(),
            r.UsuarioActualizacion = @Usuario
        OUTPUT inserted.Id, s.NegocioId
        INTO @Actualizadas (ReservaId, NegocioId)
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.SedeConfiguracionNotificacion scn ON scn.SedeId = s.Id
        WHERE r.Estado IN (1, 2)
          AND COALESCE(scn.NotificacionesActivas, 1) = 1
          AND DATEADD(
                MINUTE,
                COALESCE(scn.MinutosToleranciaNoShow, 30),
                DATEADD(MINUTE, DATEDIFF(MINUTE, 0, r.HoraInicio), CAST(r.Fecha AS DATETIME2))
          ) <= @FechaHoraActual;

        DECLARE @ReservaId INT, @NegocioId INT;
        DECLARE c CURSOR LOCAL FAST_FORWARD FOR
            SELECT a.ReservaId, a.NegocioId
            FROM @Actualizadas a;

        OPEN c;
        FETCH NEXT FROM c INTO @ReservaId, @NegocioId;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80);
            SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @ReservaId);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'RESERVAS',
                @Accion = N'AUTO_NOSHOW',
                @Entidad = N'Reserva',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Usuario,
                @DetalleJson = NULL;

            FETCH NEXT FROM c INTO @ReservaId, @NegocioId;
        END;

        CLOSE c;
        DEALLOCATE c;

        SELECT COUNT(1) AS TotalActualizadas FROM @Actualizadas;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
