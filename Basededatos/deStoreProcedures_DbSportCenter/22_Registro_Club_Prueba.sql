-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Registro directo de club con activacion automatica de prueba de 1 mes.
-- Firma:         Codex - 27/03/2026
-- =============================================

IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NegociosSuscripcion
    (
        Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        NegocioId INT NOT NULL,
        EstadoSuscripcion INT NOT NULL, -- 1 Prueba, 2 Activa, 3 Vencida, 4 Suspendida
        EsPrueba BIT NOT NULL,
        FechaInicioPrueba DATE NULL,
        FechaFinPrueba DATE NULL,
        FechaInicioPlan DATE NULL,
        FechaFinPlan DATE NULL,
        FechaCreacion DATETIME2 NOT NULL CONSTRAINT DF_NegociosSuscripcion_FechaCreacion DEFAULT (SYSUTCDATETIME()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2 NULL,
        UsuarioActualizacion NVARCHAR(200) NULL,
        CONSTRAINT FK_NegociosSuscripcion_Negocios_NegocioId FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id),
        CONSTRAINT UQ_NegociosSuscripcion_Negocio UNIQUE (NegocioId)
    );
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_RegistrarClubConPrueba
    @UsuarioId NVARCHAR(450),
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
        DECLARE @NegocioId INT;
        DECLARE @SedeId INT;
        DECLARE @CodigoSolicitud NVARCHAR(30);
        DECLARE @Secuencia INT;

        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers u WHERE u.Id = @UsuarioId)
            RAISERROR('Usuario invalido.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.UsuariosNegocio un WHERE un.UsuarioId = @UsuarioId AND un.Activo = 1)
            RAISERROR('El usuario ya tiene un negocio asociado. Solo se permite el alta inicial.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Negocios (NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro)
        VALUES (@NombreClub, NULL, NULL, 1, SYSUTCDATETIME());
        SET @NegocioId = SCOPE_IDENTITY();

        INSERT INTO dbo.Sedes (NegocioId, Nombre, Direccion, Telefono, Activo, FechaCreacion, UsuarioCreacion)
        VALUES
        (
            @NegocioId,
            CONCAT(@NombreClub, N' - Principal'),
            CONCAT(@Pais, N', ', @ProvinciaEstado, N', ', @Ciudad, N' - ', @Direccion),
            @Telefono,
            1,
            SYSUTCDATETIME(),
            @Correo
        );
        SET @SedeId = SCOPE_IDENTITY();

        INSERT INTO dbo.UsuariosNegocio (UsuarioId, NegocioId, RolNegocio, Activo)
        VALUES (@UsuarioId, @NegocioId, 1, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
        BEGIN
            INSERT INTO dbo.NegociosSuscripcion
            (
                NegocioId, EstadoSuscripcion, EsPrueba,
                FechaInicioPrueba, FechaFinPrueba,
                FechaInicioPlan, FechaFinPlan,
                FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, 1, 1,
                CAST(SYSUTCDATETIME() AS DATE),
                DATEADD(DAY, 30, CAST(SYSUTCDATETIME() AS DATE)),
                NULL, NULL,
                SYSUTCDATETIME(),
                @Correo
            );
        END;

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
                    @SedeId, 1, 90, 30, @Correo, NULL, 0, SYSUTCDATETIME(), @Correo
                );
            END;
        END;

        IF OBJECT_ID(N'dbo.SolicitudesAltaClub', N'U') IS NOT NULL
        BEGIN
            SELECT @Secuencia = COUNT(1) + 1
            FROM dbo.SolicitudesAltaClub
            WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

            SET @CodigoSolicitud = CONCAT(
                N'CLUB-',
                CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112),
                N'-',
                RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4)
            );

            INSERT INTO dbo.SolicitudesAltaClub
            (
                CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
                Pais, ProvinciaEstado, Ciudad, Direccion, Estado, ComentarioGestion,
                NegocioId, SedeId, FechaRegistro, FechaGestion, UsuarioGestion
            )
            VALUES
            (
                @CodigoSolicitud, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
                @Pais, @ProvinciaEstado, @Ciudad, @Direccion, 2, N'Autoaprobada por registro directo.',
                @NegocioId, @SedeId, SYSUTCDATETIME(), SYSUTCDATETIME(), @Correo
            );
        END;
        ELSE
        BEGIN
            SET @CodigoSolicitud = CONCAT(N'ALTA-', CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112), N'-', CONVERT(NVARCHAR(20), @NegocioId));
        END;

        IF OBJECT_ID(N'dbo.Sp_Auditoria_Registrar', N'P') IS NOT NULL
        BEGIN
            DECLARE @EntidadIdAudit NVARCHAR(80) = CONVERT(NVARCHAR(80), @NegocioId);
            EXEC dbo.Sp_Auditoria_Registrar
                @NegocioId = @NegocioId,
                @Modulo = N'ALTAS_CLUBES',
                @Accion = N'CREATE',
                @Entidad = N'Negocio',
                @EntidadId = @EntidadIdAudit,
                @Usuario = @Correo,
                @DetalleJson = NULL;
        END;

        COMMIT TRANSACTION;
        SELECT @CodigoSolicitud AS CodigoRegistro;
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
