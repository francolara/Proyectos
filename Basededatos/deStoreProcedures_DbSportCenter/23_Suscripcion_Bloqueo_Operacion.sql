-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/03/2026
-- Description:   Bloqueo operativo por suscripcion vencida/suspendida y activacion de plan.
-- Firma:         Codex - 27/03/2026
-- =============================================

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

        SELECT
            @UsuarioNegocioId = un.Id,
            @RolNegocio = un.RolNegocio,
            @NegocioNombre = n.NombreComercial
        FROM dbo.UsuariosNegocio un
        INNER JOIN dbo.Negocios n ON n.Id = un.NegocioId
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1
          AND n.Activo = 1;

        IF @UsuarioNegocioId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, N'', @ModuloCodigo, N'', N'', CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Usuario sin acceso al negocio';
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
                    N'La suscripcion del negocio esta vencida o suspendida. Activa un plan para continuar operando.';
                RETURN;
            END;
        END;

        SELECT @ModuloId = m.Id, @ModuloNombre = m.Nombre
        FROM dbo.ModulosSistema m
        WHERE m.Codigo = @ModuloCodigo AND m.Activo = 1;

        IF @ModuloId IS NULL
        BEGIN
            SELECT CAST(0 AS BIT), @NegocioId, @NegocioNombre, @ModuloCodigo, N'', CAST(@RolNegocio AS NVARCHAR(20)), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), CAST(0 AS BIT), N'Modulo no configurado';
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
            CAST(NULL AS NVARCHAR(200)) AS Mensaje;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Panel_ListarModulosPermitidos
    @UsuarioId NVARCHAR(450),
    @NegocioId INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @UsuarioNegocioId INT, @RolNegocio INT;
        DECLARE @EstadoSuscripcion INT, @EsPrueba BIT, @FechaFinPrueba DATE, @FechaFinPlan DATE;
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);

        SELECT @UsuarioNegocioId = un.Id, @RolNegocio = un.RolNegocio
        FROM dbo.UsuariosNegocio un
        WHERE un.UsuarioId = @UsuarioId
          AND un.NegocioId = @NegocioId
          AND un.Activo = 1;

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
                    m.Id,
                    m.Codigo,
                    m.Nombre,
                    CAST(0 AS BIT) AS PuedeVer,
                    CAST(0 AS BIT) AS PuedeCrear,
                    CAST(0 AS BIT) AS PuedeEditar,
                    CAST(0 AS BIT) AS PuedeEliminar
                FROM dbo.ModulosSistema m
                WHERE 1 = 0;
                RETURN;
            END;
        END;

        SELECT
            m.Id,
            m.Codigo,
            m.Nombre,
            COALESCE(up.PuedeVer, rp.PuedeVer) AS PuedeVer,
            COALESCE(up.PuedeCrear, rp.PuedeCrear) AS PuedeCrear,
            COALESCE(up.PuedeEditar, rp.PuedeEditar) AS PuedeEditar,
            COALESCE(up.PuedeEliminar, rp.PuedeEliminar) AS PuedeEliminar
        FROM dbo.ModulosSistema m
        INNER JOIN dbo.RolesNegocioPermiso rp ON rp.ModuloSistemaId = m.Id AND rp.RolNegocio = @RolNegocio
        LEFT JOIN dbo.UsuariosNegocioPermiso up ON up.ModuloSistemaId = m.Id AND up.UsuarioNegocioId = @UsuarioNegocioId
        WHERE m.Activo = 1
        ORDER BY m.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_ActivarPlan
    @NegocioId INT,
    @DiasVigencia INT = 30,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @DiasVigencia IS NULL OR @DiasVigencia <= 0
            SET @DiasVigencia = 30;

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        IF OBJECT_ID(N'dbo.NegociosSuscripcion', N'U') IS NULL
            RAISERROR('No existe la tabla NegociosSuscripcion. Ejecuta primero 22_Registro_Club_Prueba.sql.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.NegociosSuscripcion WHERE NegocioId = @NegocioId)
        BEGIN
            UPDATE dbo.NegociosSuscripcion
            SET EstadoSuscripcion = 2,
                EsPrueba = 0,
                FechaInicioPlan = CAST(SYSUTCDATETIME() AS DATE),
                FechaFinPlan = DATEADD(DAY, @DiasVigencia, CAST(SYSUTCDATETIME() AS DATE)),
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
            WHERE NegocioId = @NegocioId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.NegociosSuscripcion
            (
                NegocioId, EstadoSuscripcion, EsPrueba, FechaInicioPrueba, FechaFinPrueba,
                FechaInicioPlan, FechaFinPlan, FechaCreacion, UsuarioCreacion
            )
            VALUES
            (
                @NegocioId, 2, 0, NULL, NULL,
                CAST(SYSUTCDATETIME() AS DATE),
                DATEADD(DAY, @DiasVigencia, CAST(SYSUTCDATETIME() AS DATE)),
                SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
            );
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
