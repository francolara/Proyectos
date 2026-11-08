-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Entrega inicial HYT, esquema y catalogos para ejecucion manual.
-- =============================================
-- Revisar LEEME_DESPLIEGUE.md. Este archivo no se ejecuta automaticamente.
-- Preparado para una base nueva SistemaWebHYT, no para migrar otros esquemas.
-- Firma: FRANCO LARA | Create date: 05/11/2026 | Código SAP opcional en el maestro de clientes.
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Nombres completos del tipo de servicio en presupuesto y ubigeo de HYT. Guía como documento sin estados. Facturado admite OC completada con guía y confirmación vigente.

-- Fuente: Script/001_CrearBaseDatos.sql
USE [master];
GO
SET NOCOUNT ON;

-- CREATE DATABASE debe ejecutarse fuera de una transaccion explicita.
IF DB_ID(N'SistemaWebHYT') IS NULL
BEGIN
    EXEC(N'CREATE DATABASE [SistemaWebHYT];');
    PRINT N'Base SistemaWebHYT creada.';
END
ELSE
BEGIN
    PRINT N'La base SistemaWebHYT ya existe. No se realizaron cambios.';
END;
GO


-- Fuente: Tablas/Moneda.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Moneda', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Moneda
    (
        Codigo NVARCHAR(3) NOT NULL CONSTRAINT PK_Moneda PRIMARY KEY,
        Nombre NVARCHAR(40) NOT NULL,
        Simbolo NVARCHAR(5) NOT NULL
    );
END;
GO


-- Fuente: Tablas/TipoDocumento.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TipoDocumento', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TipoDocumento
    (
        Codigo NVARCHAR(2) NOT NULL,
        Nombre NVARCHAR(40) NOT NULL,
        CONSTRAINT PK_TipoDocumento PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoDocumento_Codigo CHECK (Codigo IN (N'01', N'03'))
    );
END;
GO


-- Fuente: Tablas/TipoIGV.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;

IF OBJECT_ID(N'dbo.TipoIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TipoIGV
    (
        Codigo NVARCHAR(10) NOT NULL,
        Nombre NVARCHAR(20) NOT NULL,
        IncluyeIGV BIT NOT NULL,
        CONSTRAINT PK_TipoIGV PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoIGV_Tratamiento CHECK
        (
            (Codigo = N'CON_IGV' AND IncluyeIGV = 1)
            OR (Codigo = N'SIN_IGV' AND IncluyeIGV = 0)
        )
    );
END;
GO


-- Fuente: Tablas/TasaIGV.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TasaIGV', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TasaIGV
    (
        FechaDesde DATE NOT NULL,
        FechaHasta DATE NULL,
        Valor DECIMAL(9,6) NOT NULL,
        CONSTRAINT PK_TasaIGV PRIMARY KEY (FechaDesde),
        CONSTRAINT UQ_TasaIGV_Valor UNIQUE (FechaDesde, Valor),
        CONSTRAINT CK_TasaIGV_Valor CHECK (Valor >= 0 AND Valor <= 1),
        CONSTRAINT CK_TasaIGV_Vigencia CHECK (FechaHasta IS NULL OR FechaHasta > FechaDesde)
    );
END;
GO


-- Fuente: Tablas/AspNetUsers.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUsers', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUsers
    (
        Id NVARCHAR(128) NOT NULL CONSTRAINT PK_AspNetUsers PRIMARY KEY,
        UserName NVARCHAR(256) NULL,
        NormalizedUserName NVARCHAR(256) NULL,
        Email NVARCHAR(256) NULL,
        NormalizedEmail NVARCHAR(256) NULL,
        EmailConfirmed BIT NOT NULL DEFAULT 0,
        PasswordHash NVARCHAR(MAX) NULL,
        SecurityStamp NVARCHAR(MAX) NULL,
        ConcurrencyStamp NVARCHAR(MAX) NULL,
        PhoneNumber NVARCHAR(MAX) NULL,
        PhoneNumberConfirmed BIT NOT NULL DEFAULT 0,
        TwoFactorEnabled BIT NOT NULL DEFAULT 0,
        LockoutEnd DATETIMEOFFSET NULL,
        LockoutEnabled BIT NOT NULL DEFAULT 1,
        AccessFailedCount INT NOT NULL DEFAULT 0,
        Nombres NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        FechaActualizacion DATETIME2(3) NULL
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetUsers') AND name = N'UserNameIndex')
    CREATE UNIQUE INDEX UserNameIndex ON dbo.AspNetUsers(NormalizedUserName) WHERE NormalizedUserName IS NOT NULL;
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetUsers') AND name = N'EmailIndex')
    CREATE INDEX EmailIndex ON dbo.AspNetUsers(NormalizedEmail);
GO


-- Fuente: Tablas/AspNetRoles.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetRoles', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetRoles
    (
        Id NVARCHAR(128) NOT NULL CONSTRAINT PK_AspNetRoles PRIMARY KEY,
        Name NVARCHAR(256) NULL,
        NormalizedName NVARCHAR(256) NULL,
        ConcurrencyStamp NVARCHAR(MAX) NULL
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.AspNetRoles') AND name = N'RoleNameIndex')
    CREATE UNIQUE INDEX RoleNameIndex ON dbo.AspNetRoles(NormalizedName) WHERE NormalizedName IS NOT NULL;
GO


-- Fuente: Tablas/AspNetUserRoles.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserRoles', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserRoles
    (
        UserId NVARCHAR(128) NOT NULL,
        RoleId NVARCHAR(128) NOT NULL,
        CONSTRAINT PK_AspNetUserRoles PRIMARY KEY (UserId, RoleId),
        CONSTRAINT FK_UserRoles_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_UserRoles_Role FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id)
    );
END;
GO


-- Fuente: Tablas/AspNetUserClaims.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserClaims', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserClaims
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetUserClaims PRIMARY KEY,
        UserId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetUserClaims_Owner FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/AspNetRoleClaims.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetRoleClaims', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetRoleClaims
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetRoleClaims PRIMARY KEY,
        RoleId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetRoleClaims_Owner FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id)
    );
END;
GO


-- Fuente: Tablas/AspNetUserLogins.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserLogins', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserLogins
    (
        LoginProvider NVARCHAR(128) NOT NULL,
        ProviderKey NVARCHAR(128) NOT NULL,
        ProviderDisplayName NVARCHAR(MAX) NULL,
        UserId NVARCHAR(128) NOT NULL,
        CONSTRAINT PK_AspNetUserLogins PRIMARY KEY (LoginProvider, ProviderKey),
        CONSTRAINT FK_UserLogins_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/AspNetUserTokens.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.AspNetUserTokens', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.AspNetUserTokens
    (
        UserId NVARCHAR(128) NOT NULL,
        LoginProvider NVARCHAR(128) NOT NULL,
        Name NVARCHAR(128) NOT NULL,
        Value NVARCHAR(MAX) NULL,
        CONSTRAINT PK_AspNetUserTokens PRIMARY KEY (UserId, LoginProvider, Name),
        CONSTRAINT FK_UserTokens_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Permiso.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Permiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Permiso
    (
        Codigo NVARCHAR(80) NOT NULL CONSTRAINT PK_Permiso PRIMARY KEY,
        Descripcion NVARCHAR(200) NOT NULL
    );
END;
GO


-- Fuente: Tablas/RolPermiso.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.RolPermiso', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.RolPermiso
    (
        RoleId NVARCHAR(128) NOT NULL,
        CodigoPermiso NVARCHAR(80) NOT NULL,
        CONSTRAINT PK_RolPermiso PRIMARY KEY (RoleId, CodigoPermiso),
        CONSTRAINT FK_RolPermiso_Rol FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id),
        CONSTRAINT FK_RolPermiso_Permiso FOREIGN KEY (CodigoPermiso) REFERENCES dbo.Permiso(Codigo)
    );
END;
GO


-- Fuente: Tablas/EventoAuditoria.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.EventoAuditoria', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.EventoAuditoria
    (
        EventoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_EventoAuditoria PRIMARY KEY,
        CorrelacionId UNIQUEIDENTIFIER NOT NULL,
        Entidad NVARCHAR(80) NOT NULL,
        RegistroId NVARCHAR(100) NOT NULL,
        Accion NVARCHAR(60) NOT NULL,
        EstadoAnterior NVARCHAR(40) NULL,
        EstadoNuevo NVARCHAR(40) NULL,
        UsuarioId NVARCHAR(128) NOT NULL,
        FechaRegistro DATETIME2(3) NOT NULL DEFAULT SYSUTCDATETIME(),
        FechaEvento DATETIME2(3) NULL,
        Motivo NVARCHAR(1000) NULL,
        CambiosJson NVARCHAR(MAX) NULL,
        CONSTRAINT FK_EventoAuditoria_Usuario FOREIGN KEY (UsuarioId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT CK_EventoAuditoria_Json CHECK (CambiosJson IS NULL OR ISJSON(CambiosJson) = 1)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.EventoAuditoria') AND name = N'IX_EventoAuditoria_Registro')
    CREATE INDEX IX_EventoAuditoria_Registro ON dbo.EventoAuditoria(Entidad, RegistroId, FechaRegistro, EventoId);
GO


-- Fuente: Tablas/ConfiguracionHYT.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfiguracionHYT', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfiguracionHYT
    (
        Id TINYINT NOT NULL CONSTRAINT PK_ConfiguracionHYT PRIMARY KEY,
        Ruc NVARCHAR(20) NOT NULL,
        RazonSocial NVARCHAR(200) NOT NULL,
        Direccion NVARCHAR(300) NULL,
        CodigoUbigeo CHAR(6) NULL,
        Telefono NVARCHAR(50) NULL,
        Correo NVARCHAR(256) NULL,
        NombreFirmante NVARCHAR(150) NULL,
        CargoFirmante NVARCHAR(150) NULL,
        FirmaPresupuestoJson NVARCHAR(2000) NULL,
        SubtituloCotizacion NVARCHAR(2000) NULL CONSTRAINT DF_ConfiguracionHYT_Subtitulo DEFAULT(N'GASFITERÍA, EBANISTERÍA, ELECTRICIDAD, PINTURA LÁTEX Y LAQUEADOS, TRABAJOS EN DRYWALL, FALSOS CIELOS,  CARPINTERÍA METÁLICA, REMODELACIONES, DEMOLICIONES Y ABASTECIMIENTO DE COMBUSTIBLE / R.P.C: 991346699 E-mail: hytserviciosg@gmail.com - rene_torres35@hotmail.com'),
        CONSTRAINT CK_ConfiguracionHYT_Unica CHECK (Id = 1),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ConfiguracionHYT_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ConfiguracionHYT_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ConfiguracionHYT_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/UnidadMedida.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.UnidadMedida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.UnidadMedida
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_UnidadMedida PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_UnidadMedida_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_UnidadMedida_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_UnidadMedida_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_UnidadMedida_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Zona.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Zona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Zona
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_Zona PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_Zona_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Zona_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Zona_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Zona_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Especialidad.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Especialidad', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Especialidad
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_Especialidad PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_Especialidad_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Especialidad_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Especialidad_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Especialidad_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/TipoServicio.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TipoServicio', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TipoServicio
    (
        Id INT IDENTITY NOT NULL CONSTRAINT PK_TipoServicio PRIMARY KEY,
        Codigo NVARCHAR(30) NOT NULL CONSTRAINT UQ_TipoServicio_Codigo UNIQUE,
        Descripcion NVARCHAR(150) NOT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_TipoServicio_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_TipoServicio_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_TipoServicio_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Cliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Cliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Cliente
    (
        ClienteId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Cliente PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Cliente_Codigo UNIQUE,
        CodigoSAP NVARCHAR(60) NULL,
        TipoEntidad NVARCHAR(20) NOT NULL,
        TipoIdentificacion NVARCHAR(20) NULL,
        NumeroIdentificacion NVARCHAR(30) NULL,
        RazonSocial NVARCHAR(200) NULL,
        Nombres NVARCHAR(150) NULL,
        Apellidos NVARCHAR(150) NULL,
        DireccionFiscal NVARCHAR(300) NULL,
        UrlCargaDocumentos NVARCHAR(2048) NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT CK_Cliente_Entidad CHECK (TipoEntidad IN (N'JURIDICA', N'NATURAL')),
        CONSTRAINT CK_Cliente_Nombre CHECK ((TipoEntidad = N'JURIDICA' AND NULLIF(LTRIM(RTRIM(RazonSocial)), N'') IS NOT NULL) OR
            (TipoEntidad = N'NATURAL' AND NULLIF(LTRIM(RTRIM(Nombres)), N'') IS NOT NULL AND NULLIF(LTRIM(RTRIM(Apellidos)), N'') IS NOT NULL)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Cliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Cliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Cliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Tienda.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Tienda', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tienda
    (
        TiendaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Tienda PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        Codigo NVARCHAR(40) NOT NULL,
        Nombre NVARCHAR(200) NOT NULL,
        Direccion NVARCHAR(300) NULL,
        CentroSAPFICO NVARCHAR(50) NULL,
        ZonaId INT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT UQ_Tienda_ClienteCodigo UNIQUE (ClienteId, Codigo),
        CONSTRAINT UQ_Tienda_ClienteId UNIQUE (ClienteId, TiendaId),
        CONSTRAINT FK_Tienda_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_Tienda_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Tienda_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Tienda_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Tienda_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/ContactoCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ContactoCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ContactoCliente
    (
        ContactoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ContactoCliente PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        TiendaId BIGINT NULL,
        Nombre NVARCHAR(150) NOT NULL,
        Telefono NVARCHAR(50) NULL,
        Cargo NVARCHAR(100) NULL,
        Correo NVARCHAR(256) NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Contacto_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT FK_Contacto_Tienda FOREIGN KEY (ClienteId, TiendaId) REFERENCES dbo.Tienda(ClienteId, TiendaId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ContactoCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ContactoCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ContactoCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/ProveedorCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ProveedorCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ProveedorCliente
    (
        Id BIGINT IDENTITY NOT NULL CONSTRAINT PK_ProveedorCliente PRIMARY KEY,
        ClienteId BIGINT NOT NULL,
        CodigoSAPProveedor NVARCHAR(60) NULL,
        ContratoMarco NVARCHAR(60) NULL,
        VigenciaDesde DATE NULL,
        VigenciaHasta DATE NULL,
        CONSTRAINT FK_ProveedorCliente_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        CONSTRAINT CK_ProveedorCliente_Fechas CHECK (VigenciaHasta IS NULL OR VigenciaDesde IS NULL OR VigenciaHasta >= VigenciaDesde),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ProveedorCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ProveedorCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ProveedorCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/ServicioMaterial.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ServicioMaterial', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ServicioMaterial
    (
        MaterialId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ServicioMaterial PRIMARY KEY,
        Posicion NVARCHAR(60) NULL,
        CodigoInterno NVARCHAR(60) NOT NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        UnidadMedidaId INT NOT NULL,
        EspecialidadId INT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Material_Unidad FOREIGN KEY (UnidadMedidaId) REFERENCES dbo.UnidadMedida(Id),
        CONSTRAINT FK_Material_Especialidad FOREIGN KEY (EspecialidadId) REFERENCES dbo.Especialidad(Id),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ServicioMaterial_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ServicioMaterial_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ServicioMaterial_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/MaterialZona.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.MaterialZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.MaterialZona
    (
        MaterialId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_MaterialZona PRIMARY KEY (MaterialId, ZonaId),
        CONSTRAINT FK_MaterialZona_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_MaterialZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO


-- Fuente: Tablas/Tarifario.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Tarifario', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Tarifario
    (
        TarifarioId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Tarifario PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Tarifario_Codigo UNIQUE,
        EsTarifarioMaterial BIT NOT NULL DEFAULT(0),
        Nombre NVARCHAR(200) NOT NULL,
        ClienteId BIGINT NULL,
        Activo BIT NOT NULL DEFAULT 1,
        CONSTRAINT FK_Tarifario_Cliente FOREIGN KEY (ClienteId) REFERENCES dbo.Cliente(ClienteId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Tarifario_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Tarifario_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Tarifario_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/TarifarioVersion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioVersion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioVersion
    (
        VersionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_TarifarioVersion PRIMARY KEY,
        TarifarioId BIGINT NOT NULL,
        NumeroVersion INT NOT NULL,
        VigenciaDesde DATE NULL,
        VigenciaHasta DATE NULL,
        CodigoMonedaPredeterminada NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGVPredeterminado NVARCHAR(10) NOT NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'BORRADOR',
        CONSTRAINT UQ_TarifarioVersion_Numero UNIQUE (TarifarioId, NumeroVersion),
        CONSTRAINT FK_TarifarioVersion_Tarifario FOREIGN KEY (TarifarioId) REFERENCES dbo.Tarifario(TarifarioId),
        CONSTRAINT FK_TarifarioVersion_Moneda FOREIGN KEY (CodigoMonedaPredeterminada) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_TarifarioVersion_IGV FOREIGN KEY (CodigoTipoIGVPredeterminado) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT CK_TarifarioVersion_Estado CHECK (Estado IN (N'BORRADOR', N'VIGENTE', N'VENCIDO', N'INACTIVO')),
        CONSTRAINT CK_TarifarioVersion_Fechas CHECK (VigenciaHasta IS NULL OR VigenciaDesde IS NULL OR VigenciaHasta >= VigenciaDesde),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_TarifarioVersion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_TarifarioVersion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_TarifarioVersion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/TarifarioVersionZona.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioVersionZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioVersionZona
    (
        VersionId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioVersionZona PRIMARY KEY (VersionId, ZonaId),
        CONSTRAINT FK_TarifarioVersionZona_Version FOREIGN KEY (VersionId) REFERENCES dbo.TarifarioVersion(VersionId),
        CONSTRAINT FK_TarifarioVersionZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO


-- Fuente: Tablas/TarifarioPartida.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_TarifarioPartida PRIMARY KEY,
        VersionId BIGINT NOT NULL,
        MaterialId BIGINT NOT NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGV NVARCHAR(10) NOT NULL,
        PrecioPEN DECIMAL(19,6) NULL,
        PrecioUSD DECIMAL(19,6) NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        TipoServicioId INT NULL,
        NumeroNecesidad NVARCHAR(60) NULL,
        CONSTRAINT FK_TarifarioPartida_Version FOREIGN KEY (VersionId) REFERENCES dbo.TarifarioVersion(VersionId),
        CONSTRAINT FK_TarifarioPartida_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_TarifarioPartida_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_TarifarioPartida_IGV FOREIGN KEY (CodigoTipoIGV) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT FK_TarifarioPartida_Tipo FOREIGN KEY (TipoServicioId) REFERENCES dbo.TipoServicio(Id),
        CONSTRAINT CK_TarifarioPartida_Precios CHECK ((PrecioPEN IS NULL OR PrecioPEN >= 0) AND (PrecioUSD IS NULL OR PrecioUSD >= 0)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_TarifarioPartida_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_TarifarioPartida_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_TarifarioPartida_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/TarifarioPartidaZona.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.TarifarioPartidaZona', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.TarifarioPartidaZona
    (
        PartidaId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioPartidaZona PRIMARY KEY (PartidaId, ZonaId),
        CONSTRAINT FK_TarifarioPartidaZona_Partida FOREIGN KEY (PartidaId) REFERENCES dbo.TarifarioPartida(PartidaId),
        CONSTRAINT FK_TarifarioPartidaZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
    );
END;
GO


-- Fuente: Tablas/Presupuesto.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Presupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Presupuesto
    (
        PresupuestoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Presupuesto PRIMARY KEY,
        NumeroInterno NVARCHAR(40) NOT NULL CONSTRAINT UQ_Presupuesto_Numero UNIQUE,
        ClienteId BIGINT NOT NULL,
        TiendaId BIGINT NOT NULL,
        TipoAtencion NVARCHAR(20) NOT NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        FechaSolicitud DATETIME2(3) NULL,
        CONSTRAINT FK_Presupuesto_Tienda FOREIGN KEY (ClienteId, TiendaId) REFERENCES dbo.Tienda(ClienteId, TiendaId),
        CONSTRAINT CK_Presupuesto_Tipo CHECK (TipoAtencion IN (N'INVITACION', N'EMERGENCIA')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Presupuesto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Presupuesto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Presupuesto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/PresupuestoRevision.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoRevision', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoRevision
    (
        RevisionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PresupuestoRevision PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL,
        NumeroRevision INT NOT NULL,
        FechaEmision DATE NOT NULL,
        NumeroCotizacionCliente NVARCHAR(60) NULL,
        ClienteNombre NVARCHAR(200) NOT NULL,
        ClienteDocumento NVARCHAR(30) NULL,
        ClienteDireccion NVARCHAR(300) NULL,
        TiendaNombre NVARCHAR(200) NOT NULL,
        TiendaDireccion NVARCHAR(300) NULL,
        FechaEjecucion DATE NULL,
        Trimestre NVARCHAR(40) NULL,
        VigenciaTexto NVARCHAR(100) NULL,
        CodigoSAPProveedor NVARCHAR(60) NULL,
        ContratoMarco NVARCHAR(60) NULL,
        Especialidad NVARCHAR(150) NULL,
        Coordinador NVARCHAR(150) NULL,
        Encargado NVARCHAR(150) NULL,
        Consideraciones NVARCHAR(MAX) NULL,
        Firmante NVARCHAR(150) NULL,
        CargoFirmante NVARCHAR(150) NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        CodigoTipoIGV NVARCHAR(10) NOT NULL,
        FechaDesdeTasaIGV DATE NOT NULL,
        TasaIGVAplicada DECIMAL(9,6) NOT NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'BORRADOR',
        Subtotal DECIMAL(19,2) NULL,
        Impuesto DECIMAL(19,2) NULL,
        Total DECIMAL(19,2) NULL,
        CONSTRAINT UQ_PresupuestoRevision_Numero UNIQUE (PresupuestoId, NumeroRevision),
        CONSTRAINT UQ_PresupuestoRevision_Vinculo UNIQUE (PresupuestoId, RevisionId),
        CONSTRAINT FK_Revision_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT FK_Revision_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_Revision_TipoIGV FOREIGN KEY (CodigoTipoIGV) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT FK_Revision_Tasa FOREIGN KEY (FechaDesdeTasaIGV, TasaIGVAplicada) REFERENCES dbo.TasaIGV(FechaDesde, Valor),
        CONSTRAINT CK_Revision_Tasa CHECK (TasaIGVAplicada >= 0 AND TasaIGVAplicada <= 1),
        CONSTRAINT CK_Revision_Estado CHECK (Estado IN (N'BORRADOR', N'ENVIADO', N'OBSERVADO', N'APROBADO', N'RECHAZADO', N'VENCIDO', N'CANCELADO')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_PresupuestoRevision_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_PresupuestoRevision_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_PresupuestoRevision_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/PresupuestoPartida.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.PresupuestoPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PresupuestoPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_PresupuestoPartida PRIMARY KEY,
        RevisionId BIGINT NOT NULL,
        NumeroLinea INT NOT NULL,
        Grupo NVARCHAR(200) NULL,
        MaterialId BIGINT NULL,
        PartidaTarifarioId BIGINT NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoInterno NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        CentroSAPFICO NVARCHAR(50) NULL,
        NumeroNecesidad NVARCHAR(60) NULL,
        TipoServicio NVARCHAR(150) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        Unidad NVARCHAR(30) NOT NULL,
        Cantidad DECIMAL(19,6) NOT NULL,
        PrecioOrigen DECIMAL(19,6) NULL,
        CodigoMonedaOrigen NVARCHAR(3) NULL,
        CodigoTipoIGVOrigen NVARCHAR(10) NULL,
        PrecioUnitario DECIMAL(19,6) NULL,
        BaseImponible DECIMAL(19,2) NULL,
        Impuesto DECIMAL(19,2) NULL,
        Total DECIMAL(19,2) NULL,
        CONSTRAINT UQ_PresupuestoPartida_Linea UNIQUE (RevisionId, NumeroLinea),
        CONSTRAINT FK_Partida_Revision FOREIGN KEY (RevisionId) REFERENCES dbo.PresupuestoRevision(RevisionId),
        CONSTRAINT FK_Partida_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_Partida_Tarifa FOREIGN KEY (PartidaTarifarioId) REFERENCES dbo.TarifarioPartida(PartidaId),
        CONSTRAINT FK_Partida_Moneda FOREIGN KEY (CodigoMonedaOrigen) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT FK_Partida_IGV FOREIGN KEY (CodigoTipoIGVOrigen) REFERENCES dbo.TipoIGV(Codigo),
        CONSTRAINT CK_Partida_Cantidad CHECK (Cantidad > 0),
        CONSTRAINT CK_Partida_Precio CHECK (PrecioUnitario IS NULL OR PrecioUnitario >= 0)
    );
END;
GO


-- Fuente: Tablas/OrdenCompraCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.OrdenCompraCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OrdenCompraCliente
    (
        OrdenCompraId BIGINT IDENTITY NOT NULL CONSTRAINT PK_OrdenCompraCliente PRIMARY KEY,
        NumeroRegistroInterno NVARCHAR(40) NOT NULL CONSTRAINT UQ_OC_Interno UNIQUE,
        NumeroOCCliente NVARCHAR(100) NOT NULL,
        PresupuestoId BIGINT NOT NULL CONSTRAINT UQ_OC_Presupuesto UNIQUE,
        RevisionId BIGINT NOT NULL,
        FechaEmisionOC DATE NOT NULL,
        FechaRecepcionOC DATE NOT NULL,
        CodigoMoneda NVARCHAR(3) NOT NULL DEFAULT N'PEN',
        Importe DECIMAL(19,2) NULL,
        Observacion NVARCHAR(2000) NULL,
        Estado NVARCHAR(30) NOT NULL DEFAULT N'REGISTRADA',
        CONSTRAINT FK_OC_Revision FOREIGN KEY (PresupuestoId, RevisionId) REFERENCES dbo.PresupuestoRevision(PresupuestoId, RevisionId),
        CONSTRAINT FK_OC_Moneda FOREIGN KEY (CodigoMoneda) REFERENCES dbo.Moneda(Codigo),
        CONSTRAINT CK_OC_Estado CHECK (Estado IN (N'REGISTRADA', N'TRABAJO_COMPLETADO', N'CERRADA_OPERATIVAMENTE', N'CANCELADA')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_OrdenCompraCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_OrdenCompraCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_OrdenCompraCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/Ejecucion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.Ejecucion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Ejecucion
    (
        EjecucionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Ejecucion PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL,
        Responsable NVARCHAR(150) NULL,
        FechaProgramada DATETIME2(3) NULL,
        FechaInicioReal DATETIME2(3) NULL,
        FechaTerminoReal DATETIME2(3) NULL,
        Descripcion NVARCHAR(2000) NULL,
        Estado NVARCHAR(30) NOT NULL DEFAULT N'PENDIENTE',
        CONSTRAINT FK_Ejecucion_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT CK_Ejecucion_Estado CHECK (Estado IN (N'PENDIENTE', N'PROGRAMADA', N'EN_EJECUCION', N'TERMINADA', N'CANCELADA')),
        CONSTRAINT CK_Ejecucion_Fechas CHECK (FechaTerminoReal IS NULL OR FechaInicioReal IS NULL OR FechaTerminoReal >= FechaInicioReal),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_Ejecucion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_Ejecucion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_Ejecucion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/GuiaConformidad.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaConformidad', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaConformidad
    (
        GuiaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_GuiaConformidad PRIMARY KEY,
        PresupuestoId BIGINT NOT NULL CONSTRAINT UQ_Guia_Presupuesto UNIQUE,
        FechaActa DATE NOT NULL,
        NumeroActa NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(2000) NULL,
        Estado NVARCHAR(20) NOT NULL DEFAULT N'REGISTRADA',
        FechaAceptacion DATETIME2(3) NULL,
        RepresentanteCliente NVARCHAR(150) NULL,
        CONSTRAINT FK_Guia_Presupuesto FOREIGN KEY (PresupuestoId) REFERENCES dbo.Presupuesto(PresupuestoId),
        CONSTRAINT CK_Guia_Estado CHECK (Estado IN (N'REGISTRADA', N'ENVIADA', N'OBSERVADA', N'ACEPTADA', N'ANULADA')),
        CONSTRAINT CK_Guia_Aceptacion CHECK (Estado <> N'ACEPTADA' OR (FechaAceptacion IS NOT NULL AND RepresentanteCliente IS NOT NULL)),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_GuiaConformidad_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_GuiaConformidad_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_GuiaConformidad_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/GuiaPartida.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaPartida', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaPartida
    (
        PartidaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_GuiaPartida PRIMARY KEY,
        GuiaId BIGINT NOT NULL,
        NumeroLinea INT NOT NULL,
        Posicion NVARCHAR(60) NULL,
        CodigoMaterialSAP NVARCHAR(60) NULL,
        Descripcion NVARCHAR(1500) NOT NULL,
        Unidad NVARCHAR(30) NOT NULL,
        Cantidad DECIMAL(19,6) NOT NULL,
        CONSTRAINT UQ_GuiaPartida_Linea UNIQUE (GuiaId, NumeroLinea),
        CONSTRAINT FK_GuiaPartida_Guia FOREIGN KEY (GuiaId) REFERENCES dbo.GuiaConformidad(GuiaId),
        CONSTRAINT CK_GuiaPartida_Cantidad CHECK (Cantidad > 0)
    );
END;
GO


-- Fuente: Tablas/SeguimientoCliente.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.SeguimientoCliente', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.SeguimientoCliente
    (
        SeguimientoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_SeguimientoCliente PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        FechaEvento DATETIME2(3) NOT NULL,
        Observacion NVARCHAR(2000) NOT NULL,
        CONSTRAINT FK_Seguimiento_OC FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_SeguimientoCliente_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_SeguimientoCliente_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_SeguimientoCliente_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/ConfirmacionFacturacion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfirmacionFacturacion', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfirmacionFacturacion
    (
        ConfirmacionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ConfirmacionFacturacion PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        FechaConfirmacion DATE NOT NULL,
        NumeroConfirmacion NVARCHAR(100) NOT NULL,
        Observacion NVARCHAR(2000) NULL,
        RepresentanteCliente NVARCHAR(150) NULL,
        Vigente BIT NOT NULL DEFAULT 1,
        CONSTRAINT UQ_Confirmacion_Vinculo UNIQUE (OrdenCompraId, ConfirmacionId),
        CONSTRAINT FK_Confirmacion_OC FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ConfirmacionFacturacion_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ConfirmacionFacturacion_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ConfirmacionFacturacion_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.ConfirmacionFacturacion') AND name = N'UX_Confirmacion_Vigente')
    CREATE UNIQUE INDEX UX_Confirmacion_Vigente ON dbo.ConfirmacionFacturacion(OrdenCompraId) WHERE Vigente = 1;
GO


-- Fuente: Tablas/ReferenciaComprobanteExterno.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ReferenciaComprobanteExterno', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ReferenciaComprobanteExterno
    (
        ReferenciaId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ReferenciaComprobanteExterno PRIMARY KEY,
        OrdenCompraId BIGINT NOT NULL,
        ConfirmacionId BIGINT NOT NULL,
        FechaDocumento DATE NOT NULL,
        CodigoTipoDocumento NVARCHAR(2) NOT NULL,
        Serie NVARCHAR(30) NOT NULL,
        NumeroDocumento NVARCHAR(50) NOT NULL,
        Vigente BIT NOT NULL DEFAULT 1,
        MotivoRetiro NVARCHAR(1000) NULL,
        CONSTRAINT FK_Referencia_Confirmacion FOREIGN KEY (OrdenCompraId, ConfirmacionId) REFERENCES dbo.ConfirmacionFacturacion(OrdenCompraId, ConfirmacionId),
        CONSTRAINT FK_Referencia_Tipo FOREIGN KEY (CodigoTipoDocumento) REFERENCES dbo.TipoDocumento(Codigo),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ReferenciaComprobanteExterno_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ReferenciaComprobanteExterno_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ReferenciaComprobanteExterno_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.ReferenciaComprobanteExterno') AND name = N'UX_Referencia_Vigente')
    CREATE UNIQUE INDEX UX_Referencia_Vigente ON dbo.ReferenciaComprobanteExterno(OrdenCompraId) WHERE Vigente = 1;
GO


-- Fuente: Tablas/ArchivoAdjunto.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ArchivoAdjunto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ArchivoAdjunto
    (
        ArchivoId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ArchivoAdjunto PRIMARY KEY,
        NombreOriginal NVARCHAR(255) NOT NULL,
        ClaveAlmacenamiento NVARCHAR(500) NOT NULL CONSTRAINT UQ_Archivo_Clave UNIQUE,
        TipoContenido NVARCHAR(150) NOT NULL,
        TamanoBytes BIGINT NOT NULL,
        SHA256 CHAR(64) NOT NULL,
        Retirado BIT NOT NULL DEFAULT 0,
        CONSTRAINT CK_Archivo_Tamano CHECK (TamanoBytes > 0),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ArchivoAdjunto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ArchivoAdjunto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ArchivoAdjunto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/OrdenCompraArchivo.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.OrdenCompraArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.OrdenCompraArchivo
    (
        OrdenCompraId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_OrdenCompraArchivo PRIMARY KEY (OrdenCompraId, ArchivoId),
        CONSTRAINT FK_OrdenCompraArchivo_Owner FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        CONSTRAINT FK_OrdenCompraArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO


-- Fuente: Tablas/GuiaArchivo.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.GuiaArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.GuiaArchivo
    (
        GuiaId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_GuiaArchivo PRIMARY KEY (GuiaId, ArchivoId),
        CONSTRAINT FK_GuiaArchivo_Owner FOREIGN KEY (GuiaId) REFERENCES dbo.GuiaConformidad(GuiaId),
        CONSTRAINT FK_GuiaArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO


-- Fuente: Tablas/ConfirmacionArchivo.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ConfirmacionArchivo', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ConfirmacionArchivo
    (
        ConfirmacionId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_ConfirmacionArchivo PRIMARY KEY (ConfirmacionId, ArchivoId),
        CONSTRAINT FK_ConfirmacionArchivo_Owner FOREIGN KEY (ConfirmacionId) REFERENCES dbo.ConfirmacionFacturacion(ConfirmacionId),
        CONSTRAINT FK_ConfirmacionArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
    );
END;
GO


-- Fuente: Tablas/ImpresionPresupuesto.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.ImpresionPresupuesto', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.ImpresionPresupuesto
    (
        ImpresionId BIGINT IDENTITY NOT NULL CONSTRAINT PK_ImpresionPresupuesto PRIMARY KEY,
        RevisionId BIGINT NOT NULL,
        Formato NVARCHAR(10) NOT NULL,
        VersionFormato NVARCHAR(20) NOT NULL,
        CONSTRAINT FK_Impresion_Revision FOREIGN KEY (RevisionId) REFERENCES dbo.PresupuestoRevision(RevisionId),
        CONSTRAINT CK_Impresion_Formato CHECK (Formato IN (N'ANTERIOR', N'NUEVO')),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ImpresionPresupuesto_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ImpresionPresupuesto_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ImpresionPresupuesto_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
    );
END;
GO


-- Fuente: Tablas/NumeracionInterna.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF OBJECT_ID(N'dbo.NumeracionInterna', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.NumeracionInterna
    (
        TipoRegistro NVARCHAR(40) NOT NULL CONSTRAINT PK_NumeracionInterna PRIMARY KEY,
        Prefijo NVARCHAR(20) NOT NULL,
        UltimoNumero BIGINT NOT NULL DEFAULT 0,
        CONSTRAINT CK_Numeracion_Numero CHECK (UltimoNumero >= 0)
    );
END;
GO


-- Fuente: Script/002_CargarTiposDocumento.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    INSERT INTO dbo.TipoDocumento (Codigo, Nombre)
    SELECT origen.Codigo, origen.Nombre
    FROM (VALUES (N'01', N'Factura'), (N'03', N'Boleta'))
        AS origen (Codigo, Nombre)
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM dbo.TipoDocumento AS tipo WITH (UPDLOCK, HOLDLOCK)
        WHERE tipo.Codigo = origen.Codigo
    );

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: Script/003_CargarTiposIGV.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    INSERT INTO dbo.TipoIGV (Codigo, Nombre, IncluyeIGV)
    SELECT origen.Codigo, origen.Nombre, origen.IncluyeIGV
    FROM (VALUES
        (N'CON_IGV', N'CON IGV', CAST(1 AS BIT)),
        (N'SIN_IGV', N'SIN IGV', CAST(0 AS BIT))
    ) AS origen (Codigo, Nombre, IncluyeIGV)
    WHERE NOT EXISTS
    (
        SELECT 1
        FROM dbo.TipoIGV AS tipo WITH (UPDLOCK, HOLDLOCK)
        WHERE tipo.Codigo = origen.Codigo
    );

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: Script/004_CargarTasaIGV.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM dbo.TasaIGV AS tasa WITH (UPDLOCK, HOLDLOCK))
    BEGIN
        INSERT INTO dbo.TasaIGV (FechaDesde, FechaHasta, Valor)
        VALUES (CONVERT(DATE, '20261103', 112), NULL, 0.18);
    END;
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: Script/005_CargarMonedas.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    INSERT INTO dbo.Moneda (Codigo, Nombre, Simbolo)
    SELECT origen.Codigo, origen.Nombre, origen.Simbolo
    FROM (VALUES (N'PEN', N'Soles', N'S/'), (N'USD', N'Dólares', N'$')) AS origen (Codigo, Nombre, Simbolo)
    WHERE NOT EXISTS (SELECT 1 FROM dbo.Moneda AS moneda WITH (UPDLOCK, HOLDLOCK) WHERE moneda.Codigo = origen.Codigo);
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: Script/006_CargarRolesPermisos.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-ADMIN')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-ADMIN', N'Administrador', N'ADMINISTRADOR', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-COMERCIAL')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-COMERCIAL', N'Comercial', N'COMERCIAL', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-OPERACIONES')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-OPERACIONES', N'Operaciones', N'OPERACIONES', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-SEGUIMIENTO')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-SEGUIMIENTO', N'Seguimiento documental', N'SEGUIMIENTO DOCUMENTAL', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.AspNetRoles AS rol WITH (UPDLOCK, HOLDLOCK) WHERE rol.Id = N'HYT-CONSULTA')
        INSERT INTO dbo.AspNetRoles (Id, Name, NormalizedName, ConcurrencyStamp) VALUES (N'HYT-CONSULTA', N'Consulta', N'CONSULTA', CONVERT(NVARCHAR(36), NEWID()));
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CLIENTES.CONSULTAR', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CLIENTES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CLIENTES.EDITAR', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TIENDAS.CONSULTAR', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TIENDAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TIENDAS.EDITAR', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'MATERIALES.CONSULTAR', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'MATERIALES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'MATERIALES.EDITAR', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'UNIDADES.CONSULTAR', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'UNIDADES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'UNIDADES.EDITAR', N'UNIDADES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ZONAS.CONSULTAR', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ZONAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ZONAS.EDITAR', N'ZONAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TARIFARIOS.CONSULTAR', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'TARIFARIOS.EDITAR', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.CONSULTAR', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.EDITAR', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'OC_CLIENTE.CONSULTAR', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'OC_CLIENTE.EDITAR', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.CONSULTAR', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.EDITAR', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIRMACIONES.CONSULTAR', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIRMACIONES.EDITAR', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'FACTURADO.CONSULTAR', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'FACTURADO.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'FACTURADO.EDITAR', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'REPORTES.CONSULTAR', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'AUDITORIA.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'AUDITORIA.CONSULTAR', N'AUDITORIA.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'USUARIOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'USUARIOS.CONSULTAR', N'USUARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'USUARIOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'USUARIOS.EDITAR', N'USUARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIGURACION.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIGURACION.CONSULTAR', N'CONFIGURACION.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'CONFIGURACION.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'CONFIGURACION.EDITAR', N'CONFIGURACION.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'PRESUPUESTOS.APROBAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'PRESUPUESTOS.APROBAR', N'PRESUPUESTOS.APROBAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'GUIAS.ACEPTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'GUIAS.ACEPTAR', N'GUIAS.ACEPTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'DOCUMENTOS.ANULAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'DOCUMENTOS.ANULAR', N'DOCUMENTOS.ANULAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ADJUNTOS.CONSULTAR', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.Permiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.Codigo = N'ADJUNTOS.EDITAR')
        INSERT INTO dbo.Permiso (Codigo, Descripcion) VALUES (N'ADJUNTOS.EDITAR', N'ADJUNTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CLIENTES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TIENDAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'MATERIALES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'UNIDADES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'UNIDADES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ZONAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ZONAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'FACTURADO.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'AUDITORIA.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'AUDITORIA.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'USUARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'USUARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'USUARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'USUARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIGURACION.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIGURACION.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'CONFIGURACION.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'CONFIGURACION.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'PRESUPUESTOS.APROBAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'PRESUPUESTOS.APROBAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'GUIAS.ACEPTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'GUIAS.ACEPTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'DOCUMENTOS.ANULAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'DOCUMENTOS.ANULAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-ADMIN' AND permiso.CodigoPermiso = N'ADJUNTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-ADMIN', N'ADJUNTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CLIENTES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CLIENTES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TIENDAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TIENDAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'MATERIALES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'MATERIALES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'TARIFARIOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'TARIFARIOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'PRESUPUESTOS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'PRESUPUESTOS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-COMERCIAL' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-COMERCIAL', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'OC_CLIENTE.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'OC_CLIENTE.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'GUIAS.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'GUIAS.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-OPERACIONES' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-OPERACIONES', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'CONFIRMACIONES.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'CONFIRMACIONES.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'FACTURADO.EDITAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'FACTURADO.EDITAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-SEGUIMIENTO' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-SEGUIMIENTO', N'ADJUNTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'CLIENTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'TIENDAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'TIENDAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'MATERIALES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'MATERIALES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'UNIDADES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'UNIDADES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'ZONAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'ZONAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'TARIFARIOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'TARIFARIOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'PRESUPUESTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'PRESUPUESTOS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'OC_CLIENTE.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'OC_CLIENTE.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'GUIAS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'GUIAS.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'CONFIRMACIONES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'CONFIRMACIONES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'FACTURADO.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'FACTURADO.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'REPORTES.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'REPORTES.CONSULTAR');
    IF NOT EXISTS (SELECT 1 FROM dbo.RolPermiso AS permiso WITH (UPDLOCK, HOLDLOCK) WHERE permiso.RoleId = N'HYT-CONSULTA' AND permiso.CodigoPermiso = N'ADJUNTOS.CONSULTAR')
        INSERT INTO dbo.RolPermiso (RoleId, CodigoPermiso) VALUES (N'HYT-CONSULTA', N'ADJUNTOS.CONSULTAR');
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: Script/009_CargarNumeracion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
SET XACT_ABORT ON;
BEGIN TRY
    BEGIN TRANSACTION;
    INSERT INTO dbo.NumeracionInterna (TipoRegistro, Prefijo, UltimoNumero)
    SELECT origen.TipoRegistro, origen.Prefijo, 0
    FROM (VALUES (N'PRESUPUESTO', N'PPTO-'), (N'OC_CLIENTE', N'OC-')) AS origen (TipoRegistro, Prefijo)
    WHERE NOT EXISTS
    (
        SELECT 1 FROM dbo.NumeracionInterna AS numeracion WITH (UPDLOCK, HOLDLOCK)
        WHERE numeracion.TipoRegistro = origen.TipoRegistro
    );
    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
    THROW;
END CATCH;
GO


-- Fuente: StoreProcedure/spu_ObtenerNumeroInterno.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObtenerNumeroInterno
    @TipoRegistro NVARCHAR(40),
    @NumeroInterno NVARCHAR(40) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @@TRANCOUNT = 0 OR XACT_STATE() <> 1
            THROW 51011, N'La numeracion requiere la transaccion del registro.', 1;
        DECLARE @Prefijo NVARCHAR(20), @Ultimo BIGINT;
        SELECT @Prefijo = numeracion.Prefijo, @Ultimo = numeracion.UltimoNumero
        FROM dbo.NumeracionInterna AS numeracion WITH (UPDLOCK, HOLDLOCK)
        WHERE numeracion.TipoRegistro = @TipoRegistro;
        IF @Ultimo IS NULL OR @Ultimo = 9223372036854775807
            THROW 51012, N'Numeracion inexistente o agotada.', 1;
        UPDATE dbo.NumeracionInterna SET UltimoNumero = @Ultimo + 1 WHERE TipoRegistro = @TipoRegistro;
        SET @NumeroInterno = CONCAT(@Prefijo, CONVERT(NVARCHAR(20), @Ultimo + 1));
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_RegistrarAuditoria.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RegistrarAuditoria
    @UsuarioId NVARCHAR(128),
    @CorrelacionId UNIQUEIDENTIFIER,
    @Entidad NVARCHAR(80),
    @RegistroId NVARCHAR(100),
    @Accion NVARCHAR(60),
    @EstadoAnterior NVARCHAR(40) = NULL,
    @EstadoNuevo NVARCHAR(40) = NULL,
    @FechaEvento DATETIME2(3) = NULL,
    @Motivo NVARCHAR(1000) = NULL,
    @CambiosJson NVARCHAR(MAX) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @@TRANCOUNT = 0 OR XACT_STATE() <> 1
            THROW 51000, N'La auditoria debe participar en una transaccion activa y valida.', 1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1)
            THROW 51001, N'Usuario de auditoria no valido o desactivado.', 1;
        IF @CambiosJson IS NOT NULL AND ISJSON(@CambiosJson) <> 1
            THROW 51002, N'El detalle de cambios debe ser JSON valido.', 1;

        INSERT INTO dbo.EventoAuditoria
            (CorrelacionId, Entidad, RegistroId, Accion, EstadoAnterior, EstadoNuevo,
             UsuarioId, FechaEvento, Motivo, CambiosJson)
        VALUES
            (@CorrelacionId, @Entidad, @RegistroId, @Accion, @EstadoAnterior, @EstadoNuevo,
             @UsuarioId, @FechaEvento, @Motivo, @CambiosJson);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_ListarClientes.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ListarClientes
    @UsuarioId NVARCHAR(128),
    @Pagina INT = 1,
    @Busqueda NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Pagina < 1 OR @Pagina > 100000000
            THROW 51003, N'Pagina fuera de rango.', 1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CLIENTES.CONSULTAR'
        )
            THROW 51004, N'No tiene permiso para consultar clientes.', 1;

        DECLARE @Patron NVARCHAR(410) = NULL;
        SET @Busqueda = NULLIF(LTRIM(RTRIM(@Busqueda)), N'');
        IF @Busqueda IS NOT NULL
            SET @Patron = N'%' + REPLACE(REPLACE(REPLACE(REPLACE(@Busqueda,
                N'~', N'~~'), N'%', N'~%'), N'_', N'~_'), N'[', N'~[') + N'%';

        -- Una materializacion por solicitud mantiene pagina y total coherentes.
        SELECT cliente.ClienteId, cliente.Codigo, cliente.TipoEntidad, cliente.NumeroIdentificacion,
            cliente.RazonSocial, cliente.Nombres, cliente.Apellidos, cliente.Activo,
            cliente.FechaRegistro, cliente.UsuarioRegistro, cliente.FechaActualizacion, cliente.UsuarioActualizacion
        INTO #Clientes
        FROM dbo.Cliente AS cliente
        WHERE @Patron IS NULL OR cliente.Codigo LIKE @Patron ESCAPE N'~'
            OR cliente.RazonSocial LIKE @Patron ESCAPE N'~'
            OR CONCAT(cliente.Nombres, N' ', cliente.Apellidos) LIKE @Patron ESCAPE N'~'
            OR cliente.NumeroIdentificacion LIKE @Patron ESCAPE N'~';

        SELECT listado.ClienteId, listado.Codigo, listado.TipoEntidad, listado.NumeroIdentificacion,
            listado.RazonSocial, listado.Nombres, listado.Apellidos, listado.Activo,
            listado.FechaRegistro, listado.UsuarioRegistro, listado.FechaActualizacion, listado.UsuarioActualizacion
        FROM #Clientes AS listado
        ORDER BY listado.Codigo, listado.ClienteId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
        SELECT COUNT_BIG(1) AS TotalRegistros FROM #Clientes AS listado;
        DROP TABLE #Clientes;
    END TRY
    BEGIN CATCH
        IF OBJECT_ID(N'tempdb..#Clientes') IS NOT NULL DROP TABLE #Clientes;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: StoreProcedure/spu_RegistrarTasaIGV.sql
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_RegistrarTasaIGV
    @UsuarioId NVARCHAR(128),
    @CorrelacionId UNIQUEIDENTIFIER,
    @FechaDesde DATE,
    @Valor DECIMAL(9,6),
    @Motivo NVARCHAR(1000)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @TransaccionPropia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0
            THROW 51005, N'Esta operacion administra su propia transaccion.', 1;
        IF @FechaDesde IS NULL OR @Valor IS NULL OR @Valor < 0 OR @Valor > 1 OR NULLIF(LTRIM(RTRIM(@Motivo)), N'') IS NULL
            THROW 51006, N'Indique fecha, tasa valida y motivo.', 1;
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS asignacion ON asignacion.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = asignacion.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CONFIGURACION.EDITAR'
        )
            THROW 51007, N'No tiene permiso para configurar tasas.', 1;

        BEGIN TRANSACTION;
        SET @TransaccionPropia = 1;
        DECLARE @UltimaFecha DATE;
        SELECT @UltimaFecha = MAX(tasa.FechaDesde) FROM dbo.TasaIGV AS tasa WITH (UPDLOCK, HOLDLOCK);
        IF @UltimaFecha IS NOT NULL AND @FechaDesde <= @UltimaFecha
            THROW 51008, N'La nueva vigencia debe ser posterior a la ultima registrada.', 1;
        UPDATE dbo.TasaIGV SET FechaHasta = @FechaDesde WHERE FechaDesde = @UltimaFecha;
        INSERT INTO dbo.TasaIGV (FechaDesde, FechaHasta, Valor) VALUES (@FechaDesde, NULL, @Valor);

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(10), @FechaDesde, 23);
        DECLARE @Detalle NVARCHAR(MAX) = (SELECT @Valor AS Valor, @FechaDesde AS FechaDesde, @UltimaFecha AS VigenciaAnterior FOR JSON PATH, WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'TasaIGV', @RegistroId,
            N'NUEVA_VIGENCIA', NULL, NULL, NULL, @Motivo, @Detalle;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @TransaccionPropia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO


-- Fuente: Script/007_ProteccionAuditoria.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
GO
CREATE OR ALTER TRIGGER dbo.tr_EventoAuditoria_Inmutable
ON dbo.EventoAuditoria
INSTEAD OF UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;
    THROW 51009, N'El historial de auditoria es inmutable.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_TasaIGV_SinSolapamiento
ON dbo.TasaIGV
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM dbo.TasaIGV AS primera WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.TasaIGV AS segunda WITH (UPDLOCK, HOLDLOCK)
            ON primera.FechaDesde < segunda.FechaDesde
            AND (primera.FechaHasta IS NULL OR segunda.FechaDesde < primera.FechaHasta)
    )
        THROW 51010, N'Las vigencias de IGV no pueden solaparse.', 1;
END;
GO

CREATE OR ALTER TRIGGER dbo.tr_ReferenciaComprobante_Requisitos
ON dbo.ReferenciaComprobanteExterno
AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS referencia
        INNER JOIN dbo.OrdenCompraCliente AS oc WITH (UPDLOCK, HOLDLOCK) ON oc.OrdenCompraId = referencia.OrdenCompraId
        INNER JOIN dbo.ConfirmacionFacturacion AS confirmacion WITH (UPDLOCK, HOLDLOCK)
            ON confirmacion.ConfirmacionId = referencia.ConfirmacionId
        WHERE referencia.Vigente = 1 AND
        (
            confirmacion.Vigente = 0 OR oc.Estado = N'CANCELADA'
            OR NOT EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS guia WITH (UPDLOCK, HOLDLOCK)
                WHERE guia.PresupuestoId = oc.PresupuestoId)
            OR confirmacion.OrdenCompraId <> oc.OrdenCompraId
            OR (oc.Estado NOT IN (N'TRABAJO_COMPLETADO', N'CERRADA_OPERATIVAMENTE') AND
                (NOT EXISTS (SELECT 1 FROM dbo.Ejecucion AS ejecucion WITH (UPDLOCK, HOLDLOCK)
                    WHERE ejecucion.PresupuestoId = oc.PresupuestoId AND ejecucion.Estado = N'TERMINADA')
                OR EXISTS (SELECT 1 FROM dbo.Ejecucion AS pendiente WITH (UPDLOCK, HOLDLOCK)
                    WHERE pendiente.PresupuestoId = oc.PresupuestoId AND pendiente.Estado NOT IN (N'TERMINADA', N'CANCELADA'))))
        )
    )
        THROW 51013, N'El marcado Facturado requiere OC valida, trabajo terminado, guia registrada y confirmacion vigente.', 1;
END;
GO
CREATE OR ALTER TRIGGER dbo.tr_Confirmacion_ProtegerFacturado
ON dbo.ConfirmacionFacturacion
AFTER UPDATE
AS
BEGIN
    SET NOCOUNT ON;
    IF EXISTS
    (
        SELECT 1 FROM inserted AS confirmacion
        INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia WITH (UPDLOCK, HOLDLOCK)
            ON referencia.ConfirmacionId = confirmacion.ConfirmacionId AND referencia.Vigente = 1
        WHERE confirmacion.Vigente = 0
    )
        THROW 51014, N'Retire primero el marcado Facturado con motivo y auditoria antes de revocar su confirmacion.', 1;
END;
GO


-- Fuente: Script/008_PermisosAplicacion.sql
USE [SistemaWebHYT];
GO
SET NOCOUNT ON;
IF DATABASE_PRINCIPAL_ID(N'hyt_app_executor') IS NULL
    EXEC(N'CREATE ROLE [hyt_app_executor] AUTHORIZATION [dbo];');
DENY INSERT, UPDATE, DELETE ON SCHEMA::dbo TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.Moneda TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TipoDocumento TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TipoIGV TO hyt_app_executor;
GRANT SELECT ON OBJECT::dbo.TasaIGV TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_ListarClientes TO hyt_app_executor;
GRANT EXECUTE ON OBJECT::dbo.spu_RegistrarTasaIGV TO hyt_app_executor;
-- No otorgar acceso directo a spu_RegistrarAuditoria ni a datos privados.
-- Los SP de operaciones futuras se concederan individualmente al implementarse.
GO

