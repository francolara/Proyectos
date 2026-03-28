-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Extension de Sp_Sedes_Crear y Sp_Sedes_Actualizar con horario/dias y fechas no laborables.
-- Firma:         Codex - 27/03/2026
-- =============================================

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
    @AtiendeLunes BIT = 1,
    @AtiendeMartes BIT = 1,
    @AtiendeMiercoles BIT = 1,
    @AtiendeJueves BIT = 1,
    @AtiendeViernes BIT = 1,
    @AtiendeSabado BIT = 1,
    @AtiendeDomingo BIT = 1,
    @HoraApertura TIME = '08:00',
    @HoraCierre TIME = '23:00',
    @FechasInhabilitadasCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraCierre <= @HoraApertura
            RAISERROR('La hora de cierre debe ser mayor a la hora de apertura.', 16, 1);
        IF COALESCE(@AtiendeLunes, 0) + COALESCE(@AtiendeMartes, 0) + COALESCE(@AtiendeMiercoles, 0) + COALESCE(@AtiendeJueves, 0) + COALESCE(@AtiendeViernes, 0) + COALESCE(@AtiendeSabado, 0) + COALESCE(@AtiendeDomingo, 0) = 0
            RAISERROR('Debes seleccionar al menos un dia de atencion.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, @Nombre, @Direccion, @Telefono, @Activo, SYSUTCDATETIME(), @Usuario);

        DECLARE @Id INT = SCOPE_IDENTITY();

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

        MERGE dbo.SedeHorarioAtencion AS tgt
        USING (SELECT @Id AS SedeId) AS src
            ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN
            UPDATE SET
                AtiendeLunes = @AtiendeLunes,
                AtiendeMartes = @AtiendeMartes,
                AtiendeMiercoles = @AtiendeMiercoles,
                AtiendeJueves = @AtiendeJueves,
                AtiendeViernes = @AtiendeViernes,
                AtiendeSabado = @AtiendeSabado,
                AtiendeDomingo = @AtiendeDomingo,
                HoraApertura = @HoraApertura,
                HoraCierre = @HoraCierre,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT
            (
                SedeId, AtiendeLunes, AtiendeMartes, AtiendeMiercoles, AtiendeJueves, AtiendeViernes, AtiendeSabado, AtiendeDomingo,
                HoraApertura, HoraCierre, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @Id, @AtiendeLunes, @AtiendeMartes, @AtiendeMiercoles, @AtiendeJueves, @AtiendeViernes, @AtiendeSabado, @AtiendeDomingo,
                @HoraApertura, @HoraCierre, SYSUTCDATETIME(), @Usuario
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

        IF @FechasInhabilitadasCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@FechasInhabilitadasCsv))) > 0
        BEGIN
            ;WITH Fechas AS
            (
                SELECT DISTINCT TRY_CONVERT(DATE, LTRIM(RTRIM(value))) AS Fecha
                FROM STRING_SPLIT(@FechasInhabilitadasCsv, N',')
                WHERE TRY_CONVERT(DATE, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeFechasInhabilitadas (SedeId, Fecha, Activo, FechaCreacion, UsuarioCreacion)
            SELECT @Id, f.Fecha, 1, SYSUTCDATETIME(), @Usuario
            FROM Fechas f;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'CREATE', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
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
    @AtiendeLunes BIT = 1,
    @AtiendeMartes BIT = 1,
    @AtiendeMiercoles BIT = 1,
    @AtiendeJueves BIT = 1,
    @AtiendeViernes BIT = 1,
    @AtiendeSabado BIT = 1,
    @AtiendeDomingo BIT = 1,
    @HoraApertura TIME = '08:00',
    @HoraCierre TIME = '23:00',
    @FechasInhabilitadasCsv NVARCHAR(MAX) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @HoraCierre <= @HoraApertura
            RAISERROR('La hora de cierre debe ser mayor a la hora de apertura.', 16, 1);
        IF COALESCE(@AtiendeLunes, 0) + COALESCE(@AtiendeMartes, 0) + COALESCE(@AtiendeMiercoles, 0) + COALESCE(@AtiendeJueves, 0) + COALESCE(@AtiendeViernes, 0) + COALESCE(@AtiendeSabado, 0) + COALESCE(@AtiendeDomingo, 0) = 0
            RAISERROR('Debes seleccionar al menos un dia de atencion.', 16, 1);

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
        USING (SELECT @Id AS SedeId) AS src ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN UPDATE SET
            NotificacionesActivas = @NotificacionesActivas,
            MinutosAnticipacionRecordatorio = @MinutosAnticipacionRecordatorio,
            MinutosToleranciaNoShow = @MinutosToleranciaNoShow,
            CorreoNotificacion = @CorreoNotificacion,
            WhatsappContacto = @WhatsappContacto,
            PermiteChatWhatsapp = @PermiteChatWhatsapp,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT (SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow, CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion)
            VALUES (@Id, @NotificacionesActivas, @MinutosAnticipacionRecordatorio, @MinutosToleranciaNoShow, @CorreoNotificacion, @WhatsappContacto, @PermiteChatWhatsapp, SYSUTCDATETIME(), @Usuario);

        MERGE dbo.SedeHorarioAtencion AS tgt
        USING (SELECT @Id AS SedeId) AS src ON tgt.SedeId = src.SedeId
        WHEN MATCHED THEN UPDATE SET
            AtiendeLunes = @AtiendeLunes, AtiendeMartes = @AtiendeMartes, AtiendeMiercoles = @AtiendeMiercoles, AtiendeJueves = @AtiendeJueves,
            AtiendeViernes = @AtiendeViernes, AtiendeSabado = @AtiendeSabado, AtiendeDomingo = @AtiendeDomingo,
            HoraApertura = @HoraApertura, HoraCierre = @HoraCierre, FechaActualizacion = SYSUTCDATETIME(), UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED THEN
            INSERT (SedeId, AtiendeLunes, AtiendeMartes, AtiendeMiercoles, AtiendeJueves, AtiendeViernes, AtiendeSabado, AtiendeDomingo, HoraApertura, HoraCierre, FechaCreacion, UsuarioCreacion)
            VALUES (@Id, @AtiendeLunes, @AtiendeMartes, @AtiendeMiercoles, @AtiendeJueves, @AtiendeViernes, @AtiendeSabado, @AtiendeDomingo, @HoraApertura, @HoraCierre, SYSUTCDATETIME(), @Usuario);

        DELETE FROM dbo.SedeServicios WHERE SedeId = @Id;
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

        DELETE FROM dbo.SedeFechasInhabilitadas WHERE SedeId = @Id;
        IF @FechasInhabilitadasCsv IS NOT NULL AND LEN(LTRIM(RTRIM(@FechasInhabilitadasCsv))) > 0
        BEGIN
            ;WITH Fechas AS
            (
                SELECT DISTINCT TRY_CONVERT(DATE, LTRIM(RTRIM(value))) AS Fecha
                FROM STRING_SPLIT(@FechasInhabilitadasCsv, N',')
                WHERE TRY_CONVERT(DATE, LTRIM(RTRIM(value))) IS NOT NULL
            )
            INSERT INTO dbo.SedeFechasInhabilitadas (SedeId, Fecha, Activo, FechaCreacion, UsuarioCreacion)
            SELECT @Id, f.Fecha, 1, SYSUTCDATETIME(), @Usuario
            FROM Fechas f;
        END;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'SEDES', @Accion = N'EDIT', @Entidad = N'Sede', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
