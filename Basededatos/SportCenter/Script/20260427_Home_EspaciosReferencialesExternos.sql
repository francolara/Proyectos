USE [DbSportCenter]
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Crea tabla HomeEspaciosReferencialesExternos y prepara soporte para union en el buscador publico del Home.
-- Firma: Codex - 27/04/2026 | Extiende referenciales externos para guardar TelefonoContacto y coordenadas (LatitudReferencia/LongitudReferencia), manteniendo script idempotente.
-- =============================================

IF OBJECT_ID('dbo.HomeEspaciosReferencialesExternos', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.HomeEspaciosReferencialesExternos
    (
        Id INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_HomeEspaciosReferencialesExternos PRIMARY KEY,
        GooglePlaceId NVARCHAR(200) NULL,
        NombreComplejo NVARCHAR(180) NOT NULL,
        NombreEspacio NVARCHAR(150) NULL,
        CodigoReferencia NVARCHAR(50) NULL,
        CodigoUbigeo CHAR(6) NOT NULL,
        TipoDeporteSuperId INT NOT NULL,
        Direccion NVARCHAR(250) NULL,
        Referencia NVARCHAR(1000) NULL,
        TelefonoContacto NVARCHAR(40) NULL,
        CorreoContacto NVARCHAR(200) NULL,
        WhatsappContacto NVARCHAR(30) NULL,
        PermiteChatWhatsapp BIT NOT NULL CONSTRAINT DF_HomeEspaciosReferencialesExternos_PermiteChatWhatsapp DEFAULT ((1)),
        TarifaReferencial DECIMAL(10,2) NULL,
        TieneIluminacion BIT NOT NULL CONSTRAINT DF_HomeEspaciosReferencialesExternos_TieneIluminacion DEFAULT ((0)),
        Techada BIT NOT NULL CONSTRAINT DF_HomeEspaciosReferencialesExternos_Techada DEFAULT ((0)),
        GoogleMapsUrl NVARCHAR(500) NULL,
        LatitudReferencia DECIMAL(10,7) NULL,
        LongitudReferencia DECIMAL(10,7) NULL,
        FotoPrincipalUrl NVARCHAR(500) NULL,
        FotosUrlsCsv NVARCHAR(MAX) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_HomeEspaciosReferencialesExternos_Activo DEFAULT ((1)),
        FechaCreacion DATETIME2(7) NOT NULL CONSTRAINT DF_HomeEspaciosReferencialesExternos_FechaCreacion DEFAULT (sysutcdatetime()),
        UsuarioCreacion NVARCHAR(200) NULL,
        FechaActualizacion DATETIME2(7) NULL,
        UsuarioActualizacion NVARCHAR(200) NULL
    );
END
GO

IF COL_LENGTH('dbo.HomeEspaciosReferencialesExternos', 'GooglePlaceId') IS NULL
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD GooglePlaceId NVARCHAR(200) NULL;
END
GO

IF COL_LENGTH('dbo.HomeEspaciosReferencialesExternos', 'TelefonoContacto') IS NULL
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD TelefonoContacto NVARCHAR(40) NULL;
END
GO

IF COL_LENGTH('dbo.HomeEspaciosReferencialesExternos', 'LatitudReferencia') IS NULL
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD LatitudReferencia DECIMAL(10,7) NULL;
END
GO

IF COL_LENGTH('dbo.HomeEspaciosReferencialesExternos', 'LongitudReferencia') IS NULL
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD LongitudReferencia DECIMAL(10,7) NULL;
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Lista tipos de deporte supermaestro para el barrido de referenciales externos.
-- Firma: Codex - 27/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ReferencialesExternos_ListarTiposDeporteSuper
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            tsm.Id,
            tsm.Nombre
        FROM dbo.TiposDeporteSuperMaestro tsm
        WHERE tsm.Activo = 1
        ORDER BY tsm.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Lista referenciales externos del Home para superadmin, con filtros y paginacion.
-- Firma: Codex - 27/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ReferencialesExternos_ListarAdmin
    @CodigoDepartamento CHAR(2) = NULL,
    @CodigoProvincia CHAR(4) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @BuscarNombre NVARCHAR(180) = NULL,
    @Pagina INT = 1,
    @TamanoPagina INT = 50,
    @SoloActivos BIT = 1,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @CodigoDepartamento = NULLIF(LTRIM(RTRIM(@CodigoDepartamento)), '');
        SET @CodigoProvincia = NULLIF(LTRIM(RTRIM(@CodigoProvincia)), '');
        SET @CodigoUbigeo = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @BuscarNombre = NULLIF(LTRIM(RTRIM(@BuscarNombre)), '');
        SET @Pagina = CASE WHEN @Pagina IS NULL OR @Pagina < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN @TamanoPagina IS NULL OR @TamanoPagina < 1 THEN 50 ELSE @TamanoPagina END;

        IF OBJECT_ID('tempdb..#BaseReferenciales') IS NOT NULL
            DROP TABLE #BaseReferenciales;

        CREATE TABLE #BaseReferenciales
        (
            Id INT NOT NULL,
            NombreComplejo NVARCHAR(180) NOT NULL,
            NombreEspacio NVARCHAR(150) NULL,
            TipoDeporte NVARCHAR(120) NOT NULL,
            Departamento NVARCHAR(120) NOT NULL,
            Provincia NVARCHAR(120) NOT NULL,
            Distrito NVARCHAR(120) NOT NULL,
            Direccion NVARCHAR(250) NULL,
            GoogleMapsUrl NVARCHAR(500) NULL,
            Activo BIT NOT NULL,
            FechaActualizacion DATETIME2(7) NULL,
            UsuarioActualizacion NVARCHAR(200) NULL
        );

        INSERT INTO #BaseReferenciales
        (
            Id,
            NombreComplejo,
            NombreEspacio,
            TipoDeporte,
            Departamento,
            Provincia,
            Distrito,
            Direccion,
            GoogleMapsUrl,
            Activo,
            FechaActualizacion,
            UsuarioActualizacion
        )
        SELECT
            he.Id,
            he.NombreComplejo,
            he.NombreEspacio,
            tsm.Nombre AS TipoDeporte,
            udp.Nombre AS Departamento,
            upp.Nombre AS Provincia,
            ud.Nombre AS Distrito,
            he.Direccion,
            he.GoogleMapsUrl,
            he.Activo,
            COALESCE(he.FechaActualizacion, he.FechaCreacion) AS FechaActualizacion,
            COALESCE(he.UsuarioActualizacion, he.UsuarioCreacion) AS UsuarioActualizacion
        FROM dbo.HomeEspaciosReferencialesExternos he
        INNER JOIN dbo.UbigeoDistritos ud ON ud.CodigoUbigeo = he.CodigoUbigeo
        INNER JOIN dbo.UbigeoProvincias upp ON upp.CodigoProvincia = ud.CodigoProvincia
        INNER JOIN dbo.UbigeoDepartamentos udp ON udp.CodigoDepartamento = ud.CodigoDepartamento
        INNER JOIN dbo.TiposDeporteSuperMaestro tsm ON tsm.Id = he.TipoDeporteSuperId
        WHERE (@SoloActivos IS NULL OR he.Activo = @SoloActivos)
          AND (@CodigoDepartamento IS NULL OR ud.CodigoDepartamento = @CodigoDepartamento)
          AND (@CodigoProvincia IS NULL OR ud.CodigoProvincia = @CodigoProvincia)
          AND (@CodigoUbigeo IS NULL OR he.CodigoUbigeo = @CodigoUbigeo)
          AND (
                @BuscarNombre IS NULL
                OR he.NombreComplejo LIKE N'%' + @BuscarNombre + N'%'
                OR ISNULL(he.NombreEspacio, N'') LIKE N'%' + @BuscarNombre + N'%'
                OR ISNULL(he.Direccion, N'') LIKE N'%' + @BuscarNombre + N'%'
              );

        SELECT @TotalRegistros = COUNT(1)
        FROM #BaseReferenciales;

        SELECT
            b.Id,
            b.NombreComplejo,
            b.NombreEspacio,
            b.TipoDeporte,
            b.Departamento,
            b.Provincia,
            b.Distrito,
            b.Direccion,
            b.GoogleMapsUrl,
            b.Activo,
            b.FechaActualizacion,
            b.UsuarioActualizacion
        FROM #BaseReferenciales b
        ORDER BY
            b.Activo DESC,
            b.NombreComplejo ASC,
            b.Id DESC
        OFFSET (@Pagina - 1) * @TamanoPagina ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   27/04/2026
-- Description:   Inactiva un referencial externo del Home desde superadmin.
-- Firma: Codex - 27/04/2026
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Home_ReferencialesExternos_Inactivar
    @Id INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Usuario = COALESCE(NULLIF(LTRIM(RTRIM(@Usuario)), ''), 'owner-platform');

        IF NOT EXISTS (SELECT 1 FROM dbo.HomeEspaciosReferencialesExternos WHERE Id = @Id)
            RAISERROR('Referencial externo no encontrado.', 16, 1);

        UPDATE dbo.HomeEspaciosReferencialesExternos
           SET Activo = 0,
               FechaActualizacion = SYSUTCDATETIME(),
               UsuarioActualizacion = @Usuario
         WHERE Id = @Id
           AND Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_HomeEspaciosReferencialesExternos_UbigeoDistritos_CodigoUbigeo'
)
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD CONSTRAINT FK_HomeEspaciosReferencialesExternos_UbigeoDistritos_CodigoUbigeo
        FOREIGN KEY (CodigoUbigeo) REFERENCES dbo.UbigeoDistritos (CodigoUbigeo);
END
GO

CREATE OR ALTER PROCEDURE dbo.Sp_Home_ReferencialExterno_UpsertDesdeGoogle
    @GooglePlaceId NVARCHAR(200),
    @NombreComplejo NVARCHAR(180),
    @NombreEspacio NVARCHAR(150) = NULL,
    @CodigoReferencia NVARCHAR(50) = NULL,
    @CodigoUbigeo CHAR(6),
    @TipoDeporteSuperId INT,
    @Direccion NVARCHAR(250) = NULL,
    @Referencia NVARCHAR(1000) = NULL,
    @TelefonoContacto NVARCHAR(40) = NULL,
    @CorreoContacto NVARCHAR(200) = NULL,
    @WhatsappContacto NVARCHAR(30) = NULL,
    @PermiteChatWhatsapp BIT = 0,
    @TarifaReferencial DECIMAL(10,2) = NULL,
    @TieneIluminacion BIT = 0,
    @Techada BIT = 0,
    @GoogleMapsUrl NVARCHAR(500) = NULL,
    @FotoPrincipalUrl NVARCHAR(500) = NULL,
    @FotosUrlsCsv NVARCHAR(MAX) = NULL,
    @LatitudReferencia DECIMAL(10,7) = NULL,
    @LongitudReferencia DECIMAL(10,7) = NULL,
    @Activo BIT = 1,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @GooglePlaceId = NULLIF(LTRIM(RTRIM(@GooglePlaceId)), '');
        SET @NombreComplejo = LTRIM(RTRIM(@NombreComplejo));
        SET @NombreEspacio = NULLIF(LTRIM(RTRIM(@NombreEspacio)), '');
        SET @CodigoReferencia = NULLIF(LTRIM(RTRIM(@CodigoReferencia)), '');
        SET @Direccion = NULLIF(LTRIM(RTRIM(@Direccion)), '');
        SET @Referencia = NULLIF(LTRIM(RTRIM(@Referencia)), '');
        SET @TelefonoContacto = NULLIF(LTRIM(RTRIM(@TelefonoContacto)), '');
        SET @CorreoContacto = NULLIF(LTRIM(RTRIM(@CorreoContacto)), '');
        SET @WhatsappContacto = NULLIF(LTRIM(RTRIM(@WhatsappContacto)), '');
        SET @GoogleMapsUrl = NULLIF(LTRIM(RTRIM(@GoogleMapsUrl)), '');
        SET @FotoPrincipalUrl = NULLIF(LTRIM(RTRIM(@FotoPrincipalUrl)), '');
        SET @FotosUrlsCsv = NULLIF(LTRIM(RTRIM(@FotosUrlsCsv)), '');
        SET @Usuario = COALESCE(NULLIF(LTRIM(RTRIM(@Usuario)), ''), 'sync-google');

        IF @GooglePlaceId IS NULL
            RAISERROR('GooglePlaceId es obligatorio para sincronizar referenciales externos.', 16, 1);

        IF @NombreComplejo = ''
            RAISERROR('NombreComplejo es obligatorio.', 16, 1);

        IF EXISTS (SELECT 1 FROM dbo.HomeEspaciosReferencialesExternos WHERE GooglePlaceId = @GooglePlaceId)
        BEGIN
            UPDATE dbo.HomeEspaciosReferencialesExternos
               SET NombreComplejo = @NombreComplejo,
                   NombreEspacio = @NombreEspacio,
                   CodigoReferencia = @CodigoReferencia,
                   CodigoUbigeo = @CodigoUbigeo,
                   TipoDeporteSuperId = @TipoDeporteSuperId,
                   Direccion = @Direccion,
                   Referencia = @Referencia,
                   TelefonoContacto = @TelefonoContacto,
                   CorreoContacto = @CorreoContacto,
                   WhatsappContacto = @WhatsappContacto,
                   PermiteChatWhatsapp = COALESCE(@PermiteChatWhatsapp, 0),
                   TarifaReferencial = @TarifaReferencial,
                   TieneIluminacion = COALESCE(@TieneIluminacion, 0),
                   Techada = COALESCE(@Techada, 0),
                   GoogleMapsUrl = @GoogleMapsUrl,
                   LatitudReferencia = @LatitudReferencia,
                   LongitudReferencia = @LongitudReferencia,
                   FotoPrincipalUrl = @FotoPrincipalUrl,
                   FotosUrlsCsv = @FotosUrlsCsv,
                   Activo = COALESCE(@Activo, 1),
                   FechaActualizacion = SYSUTCDATETIME(),
                   UsuarioActualizacion = @Usuario
             WHERE GooglePlaceId = @GooglePlaceId;

            SELECT 'ACTUALIZADO' AS Accion;
            RETURN;
        END

        INSERT INTO dbo.HomeEspaciosReferencialesExternos
        (
            GooglePlaceId,
            NombreComplejo,
            NombreEspacio,
            CodigoReferencia,
            CodigoUbigeo,
            TipoDeporteSuperId,
            Direccion,
            Referencia,
            TelefonoContacto,
            CorreoContacto,
            WhatsappContacto,
            PermiteChatWhatsapp,
            TarifaReferencial,
            TieneIluminacion,
            Techada,
            GoogleMapsUrl,
            LatitudReferencia,
            LongitudReferencia,
            FotoPrincipalUrl,
            FotosUrlsCsv,
            Activo,
            FechaCreacion,
            UsuarioCreacion
        )
        VALUES
        (
            @GooglePlaceId,
            @NombreComplejo,
            @NombreEspacio,
            @CodigoReferencia,
            @CodigoUbigeo,
            @TipoDeporteSuperId,
            @Direccion,
            @Referencia,
            @TelefonoContacto,
            @CorreoContacto,
            @WhatsappContacto,
            COALESCE(@PermiteChatWhatsapp, 0),
            @TarifaReferencial,
            COALESCE(@TieneIluminacion, 0),
            COALESCE(@Techada, 0),
            @GoogleMapsUrl,
            @LatitudReferencia,
            @LongitudReferencia,
            @FotoPrincipalUrl,
            @FotosUrlsCsv,
            COALESCE(@Activo, 1),
            SYSUTCDATETIME(),
            @Usuario
        );

        SELECT 'INSERTADO' AS Accion;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID('dbo.HomeEspaciosReferencialesExternos')
      AND name = 'UQ_HomeEspaciosReferencialesExternos_GooglePlaceId'
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX UQ_HomeEspaciosReferencialesExternos_GooglePlaceId
        ON dbo.HomeEspaciosReferencialesExternos (GooglePlaceId)
        WHERE GooglePlaceId IS NOT NULL;
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.foreign_keys
    WHERE name = 'FK_HomeEspaciosReferencialesExternos_TiposDeporteSuperMaestro_TipoDeporteSuperId'
)
BEGIN
    ALTER TABLE dbo.HomeEspaciosReferencialesExternos
    ADD CONSTRAINT FK_HomeEspaciosReferencialesExternos_TiposDeporteSuperMaestro_TipoDeporteSuperId
        FOREIGN KEY (TipoDeporteSuperId) REFERENCES dbo.TiposDeporteSuperMaestro (Id);
END
GO

IF NOT EXISTS
(
    SELECT 1
    FROM sys.indexes
    WHERE object_id = OBJECT_ID('dbo.HomeEspaciosReferencialesExternos')
      AND name = 'IX_HomeEspaciosReferencialesExternos_Busqueda'
)
BEGIN
    CREATE NONCLUSTERED INDEX IX_HomeEspaciosReferencialesExternos_Busqueda
        ON dbo.HomeEspaciosReferencialesExternos (Activo, TipoDeporteSuperId, CodigoUbigeo);
END
GO
