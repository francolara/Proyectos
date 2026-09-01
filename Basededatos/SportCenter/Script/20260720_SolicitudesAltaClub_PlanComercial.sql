
GO

-- Firma: FRANCO LARA - 20/07/2026 | Guarda el plan comercial publico y aplica sus limites internos al aprobar altas de complejos deportivos.
-- Firma: FRANCO LARA - 21/07/2026 | Renombra los planes publicos a Prueba, Esencial y Pro, hace idempotente la migracion, conserva la asignacion interna y establece 15 dias de prueba por defecto.
IF COL_LENGTH(N'dbo.SolicitudesAltaClub', N'PlanComercial') IS NULL
BEGIN
    ALTER TABLE dbo.SolicitudesAltaClub
        ADD PlanComercial NVARCHAR(20) NOT NULL
            CONSTRAINT DF_SolicitudesAltaClub_PlanComercial DEFAULT (N'PRUEBA');
END;
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.default_constraints AS dc
    INNER JOIN sys.columns AS c
        ON c.object_id = dc.parent_object_id
       AND c.column_id = dc.parent_column_id
    WHERE dc.parent_object_id = OBJECT_ID(N'dbo.SolicitudesAltaClub')
      AND c.name = N'PlanComercial'
)
BEGIN
    ALTER TABLE dbo.SolicitudesAltaClub
        ADD CONSTRAINT DF_SolicitudesAltaClub_PlanComercial
            DEFAULT (N'PRUEBA') FOR PlanComercial;
END;

IF EXISTS
(
    SELECT 1
    FROM sys.check_constraints
    WHERE parent_object_id = OBJECT_ID(N'dbo.SolicitudesAltaClub')
      AND name = N'CK_SolicitudesAltaClub_PlanComercial'
)
    ALTER TABLE dbo.SolicitudesAltaClub DROP CONSTRAINT CK_SolicitudesAltaClub_PlanComercial;
GO

UPDATE dbo.SolicitudesAltaClub
SET PlanComercial = CASE UPPER(LTRIM(RTRIM(COALESCE(PlanComercial, N''))))
                        WHEN N'ESENCIAL' THEN N'ESENCIAL'
                        WHEN N'EMPRENDEDOR' THEN N'ESENCIAL'
                        WHEN N'PRO' THEN N'PRO'
                        WHEN N'PROFESIONAL' THEN N'PRO'
                        ELSE N'PRUEBA'
                    END;
GO

ALTER TABLE dbo.SolicitudesAltaClub
    WITH CHECK ADD CONSTRAINT CK_SolicitudesAltaClub_PlanComercial
        CHECK (PlanComercial IN (N'PRUEBA', N'ESENCIAL', N'PRO'));

ALTER TABLE dbo.SolicitudesAltaClub
    CHECK CONSTRAINT CK_SolicitudesAltaClub_PlanComercial;
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
    @Direccion NVARCHAR(250),
    @PlanComercial NVARCHAR(20) = N'PRUEBA'
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Secuencia INT;
        DECLARE @Codigo NVARCHAR(30);

        SET @PlanComercial = CASE UPPER(LTRIM(RTRIM(COALESCE(@PlanComercial, N''))))
                                  WHEN N'ESENCIAL' THEN N'ESENCIAL'
                                  WHEN N'EMPRENDEDOR' THEN N'ESENCIAL'
                                  WHEN N'PRO' THEN N'PRO'
                                  WHEN N'PROFESIONAL' THEN N'PRO'
                                  ELSE N'PRUEBA'
                              END;

        SELECT @Secuencia = COUNT(1) + 1
        FROM dbo.SolicitudesAltaClub
        WHERE CAST(FechaRegistro AS DATE) = CAST(SYSUTCDATETIME() AS DATE);

        SET @Codigo = CONCAT(N'CLUB-', CONVERT(NVARCHAR(8), CAST(SYSUTCDATETIME() AS DATE), 112), N'-', RIGHT(CONCAT(N'0000', CONVERT(NVARCHAR(10), @Secuencia)), 4));

        INSERT INTO dbo.SolicitudesAltaClub
        (
            CodigoSolicitud, NombreContacto, Telefono, Correo, RelacionClub, NombreClub,
            Pais, ProvinciaEstado, Ciudad, Direccion, PlanComercial, Estado, FechaRegistro
        )
        VALUES
        (
            @Codigo, @NombreContacto, @Telefono, @Correo, @RelacionClub, @NombreClub,
            @Pais, @ProvinciaEstado, @Ciudad, @Direccion, @PlanComercial, 1, SYSUTCDATETIME()
        );

        SELECT @Codigo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Listar
    @Estado INT = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT,
    @TotalPendientes INT OUTPUT,
    @TotalAprobados INT OUTPUT,
    @TotalRechazados INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;

        SELECT @TotalPendientes = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 1;
        SELECT @TotalAprobados = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 2;
        SELECT @TotalRechazados = COUNT(1) FROM dbo.SolicitudesAltaClub WHERE Estado = 3;
        SELECT @TotalRegistros = COUNT(1) FROM dbo.SolicitudesAltaClub ac WHERE @Estado IS NULL OR ac.Estado = @Estado;

        SELECT
            ac.Id, ac.CodigoSolicitud, ac.NombreContacto, ac.Telefono, ac.Correo, ac.RelacionClub,
            ac.NombreClub, ac.Pais, ac.ProvinciaEstado, ac.Ciudad, ac.Direccion, ac.Estado,
            ac.ComentarioGestion, ac.NegocioId, ac.SedeId, ac.FechaRegistro, ac.FechaGestion,
            COALESCE(ac.PlanComercial, N'PRUEBA') AS PlanComercial
        FROM dbo.SolicitudesAltaClub ac
        WHERE @Estado IS NULL OR ac.Estado = @Estado
        ORDER BY ac.FechaRegistro DESC
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_ObtenerPorId
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            ac.Id, ac.CodigoSolicitud, ac.NombreContacto, ac.Telefono, ac.Correo, ac.RelacionClub,
            ac.NombreClub, ac.Pais, ac.ProvinciaEstado, ac.Ciudad, ac.Direccion, ac.Estado,
            ac.ComentarioGestion, ac.NegocioId, ac.SedeId, ac.FechaRegistro, ac.FechaGestion,
            COALESCE(ac.PlanComercial, N'PRUEBA') AS PlanComercial
        FROM dbo.SolicitudesAltaClub ac
        WHERE ac.Id = @Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO

CREATE OR ALTER PROCEDURE dbo.Sp_AltasClubes_Aprobar
    @Id INT,
    @Usuario NVARCHAR(200),
    @ComentarioGestion NVARCHAR(300) = NULL,
    @DiasPrueba INT = 15
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Correo NVARCHAR(200), @NombreClub NVARCHAR(200), @Telefono NVARCHAR(30), @Direccion NVARCHAR(250), @Ciudad NVARCHAR(120), @EstadoActual INT;
        DECLARE @NegocioId INT, @SedeId INT, @UsuarioId NVARCHAR(450);
        DECLARE @Hoy DATE = CAST(SYSUTCDATETIME() AS DATE);

        IF @DiasPrueba IS NULL OR @DiasPrueba <= 0
            SET @DiasPrueba = 15;

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

        INSERT INTO dbo.Negocios
        (
            NombreComercial, RazonSocial, DocumentoFiscal, Activo, FechaRegistro, MonedaId, TipoPlan
        )
        VALUES
        (
            @NombreClub, NULL, NULL, 1, SYSUTCDATETIME(), NULL, N'Basico'
        );
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
END;
GO
