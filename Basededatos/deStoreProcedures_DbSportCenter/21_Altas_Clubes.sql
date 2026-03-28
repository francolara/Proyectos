-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Sprint 8 - Alta publica "Software para Clubes" y aprobacion interna.
-- =============================================

IF OBJECT_ID(N'dbo.SolicitudesAltaClub', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SolicitudesAltaClub
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        CodigoSolicitud NVARCHAR(30) NOT NULL UNIQUE,
        NombreContacto NVARCHAR(200) NOT NULL,
        Telefono NVARCHAR(30) NOT NULL,
        Correo NVARCHAR(200) NOT NULL,
        RelacionClub NVARCHAR(80) NOT NULL,
        NombreClub NVARCHAR(200) NOT NULL,
        Pais NVARCHAR(80) NOT NULL,
        ProvinciaEstado NVARCHAR(120) NOT NULL,
        Ciudad NVARCHAR(120) NOT NULL,
        Direccion NVARCHAR(250) NOT NULL,
        Estado INT NOT NULL CONSTRAINT DF_SolicitudesAltaClub_Estado DEFAULT (1), -- 1 Pendiente, 2 Aprobada, 3 Rechazada
        ComentarioGestion NVARCHAR(300) NULL,
        NegocioId INT NULL,
        SedeId INT NULL,
        FechaRegistro DATETIME2 NOT NULL CONSTRAINT DF_SolicitudesAltaClub_FechaRegistro DEFAULT (SYSUTCDATETIME()),
        FechaGestion DATETIME2 NULL,
        UsuarioGestion NVARCHAR(200) NULL,
        CONSTRAINT FK_SolicitudesAltaClub_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id),
        CONSTRAINT FK_SolicitudesAltaClub_Sedes_SedeId FOREIGN KEY (SedeId) REFERENCES dbo.Sedes (Id)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_SolicitarAltaClub
    @NombreContacto NVARCHAR(200),
    @Telefono NVARCHAR(30),
    @Correo NVARCHAR(200),
    @RelacionClub NVARCHAR(80),
    @NombreClub NVARCHAR(200),
    @Pais NVARCHAR(80),
    @ProvinciaEstado NVARCHAR(120),
    @Ciudad NVARCHAR(120),
    @Direccion NVARCHAR(250)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Secuencia INT;
        DECLARE @Codigo NVARCHAR(30);

        SELECT @Secuencia = COUNT(1) + 1
        FROM dbo.SolicitudesAltaClub
        WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

        SET @Codigo = CONCAT(
            N'CLUB-',
            CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112),
            N'-',
            RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4)
        );

        INSERT INTO dbo.SolicitudesAltaClub
        (
            CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
            Pais, ProvinciaEstado, Ciudad, Direccion, Estado, FechaRegistro
        )
        VALUES
        (
            @Codigo, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
            @Pais, @ProvinciaEstado, @Ciudad, @Direccion, 1, SYSUTCDATETIME()
        );

        SELECT @Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Listar
    @Estado INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            ac.Id,
            ac.CodigoSolicitud,
            ac.NombreContacto,
            ac.Telefono,
            ac.Correo,
            ac.RelacionClub,
            ac.NombreClub,
            ac.Pais,
            ac.ProvinciaEstado,
            ac.Ciudad,
            ac.Direccion,
            ac.Estado,
            ac.ComentarioGestion,
            ac.NegocioId,
            ac.SedeId,
            ac.FechaRegistro,
            ac.FechaGestion
        FROM dbo.SolicitudesAltaClub ac
        WHERE (@Estado IS NULL OR ac.Estado = @Estado)
        ORDER BY ac.FechaRegistro DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Aprobar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Correo NVARCHAR(200), @NombreClub NVARCHAR(200), @Telefono NVARCHAR(30), @Direccion NVARCHAR(250), @Ciudad NVARCHAR(120), @EstadoActual INT;
        DECLARE @NegocioId INT, @SedeId INT;

        SELECT
            @Correo = ac.Correo,
            @NombreClub = ac.NombreClub,
            @Telefono = ac.Telefono,
            @Direccion = ac.Direccion,
            @Ciudad = ac.Ciudad,
            @EstadoActual = ac.Estado
        FROM dbo.SolicitudesAltaClub ac
        WHERE ac.Id = @Id;

        IF @EstadoActual IS NULL
            RAISERROR('Solicitud no encontrada.', 16, 1);

        IF @EstadoActual <> 1
            RAISERROR('Solo se pueden aprobar solicitudes pendientes.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Negocios (NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro)
        VALUES (@NombreClub, NULL, NULL, 1, SYSUTCDATETIME());
        SET @NegocioId = SCOPE_IDENTITY();

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES (@NegocioId, CONCAT(@NombreClub, N' - Principal'), CONCAT(@Ciudad, N' - ', @Direccion), @Telefono, 1, SYSUTCDATETIME(), @Usuario);
        SET @SedeId = SCOPE_IDENTITY();

        IF OBJECT_ID(N'dbo.SedeConfiguracionNotificacion', N'U') IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.SedeConfiguracionNotificacion WHERE SedeId = @SedeId)
            BEGIN
                INSERT INTO dbo.SedeConfiguracionNotificacion
                (
                    SedeId, NotificacionesActivas, MinutosAnticipacionRecordatorio, MinutosToleranciaNoShow,
                    CorreoNotificacion, WhatsappContacto, PermiteChatWhatsapp, FechaCreacion, UsuarioCreacion
                )
                VALUES
                (
                    @SedeId, 1, 90, 30, @Correo, NULL, 0, SYSUTCDATETIME(), @Usuario
                );
            END;
        END;

        DECLARE @UsuarioId NVARCHAR(450);
        SELECT TOP (1) @UsuarioId = u.Id
        FROM dbo.AspNetUsers u
        WHERE u.NormalizedEmail = UPPER(@Correo);

        IF @UsuarioId IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.UsuariosNegocio WHERE UsuarioId = @UsuarioId AND NegocioId = @NegocioId)
            BEGIN
                INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, Activo)
                VALUES (@UsuarioId, @NegocioId, 1, 1);
            END;
        END;

        UPDATE dbo.SolicitudesAltaClub
        SET Estado = 2,
            ComentarioGestion = @ComentarioGestion,
            NegocioId = @NegocioId,
            SedeId = @SedeId,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario
        WHERE Id = @Id;

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

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Rechazar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        UPDATE dbo.SolicitudesAltaClub
        SET Estado = 3,
            ComentarioGestion = @ComentarioGestion,
            FechaGestion = SYSUTCDATETIME(),
            UsuarioGestion = @Usuario
        WHERE Id = @Id
          AND Estado = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
