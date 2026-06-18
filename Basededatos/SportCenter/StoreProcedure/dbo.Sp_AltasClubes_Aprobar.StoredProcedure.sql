-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/04/2026
-- Firma:         Ajuste de aprobacion de altas con dias de prueba configurables y creacion de suscripcion inicial.
-- Firma:         FRANCO LARA - 18/06/2026 | Registra TipoPlan Basico por defecto al crear nuevos negocios desde altas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Aprobar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL,
    @DiasPrueba INT = 7
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Correo NVARCHAR(200), @NombreClub NVARCHAR(200), @Telefono NVARCHAR(30), @Direccion NVARCHAR(250), @Ciudad NVARCHAR(120), @EstadoActual INT;
        DECLARE @NegocioId INT, @SedeId INT;
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);

        IF @DiasPrueba IS NULL OR @DiasPrueba <= 0
            SET @DiasPrueba = 7;

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

        INSERT INTO dbo.Negocios (NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro, MonedaId, TipoPlan)
        VALUES (@NombreClub, NULL, NULL, 1, SYSUTCDATETIME(), NULL, N'Basico');
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

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NOT NULL
        BEGIN
            IF EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
            BEGIN
                UPDATE dbo.NegociosSuscripcion
                SET EstadoSuscripcion = 1,
                    EsPrueba = 1,
                    FechaInicioPrueba = @Hoy,
                    FechaFinPrueba = DATEADD(DAY, @DiasPrueba, @Hoy),
                    FechaInicioPlan = NULL,
                    FechaFinPlan = NULL,
                    FechaFinGracia = NULL,
                    TipoCobro = NULL,
                    DiasGracia = COALESCE(DiasGracia, 5),
                    FechaActualizacion = SYSUTCDATETIME(),
                    UsuarioActualizacion = @Usuario
                WHERE NegocioId = @NegocioId;
            END
            ELSE
            BEGIN
                INSERT INTO dbo.NegociosSuscripcion
                (
                    NegocioId, EstadoSuscripcion, EsPrueba,
                    FechaInicioPrueba, FechaFinPrueba,
                    FechaInicioPlan, FechaFinPlan,
                    TipoCobro, DiasGracia, FechaFinGracia,
                    FechaCreacion, UsuarioCreacion
                )
                VALUES
                (
                    @NegocioId, 1, 1,
                    @Hoy, DATEADD(DAY, @DiasPrueba, @Hoy),
                    NULL, NULL,
                    NULL, 5, NULL,
                    SYSUTCDATETIME(), @Usuario
                );
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
