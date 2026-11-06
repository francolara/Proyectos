# Diccionario físico de la etapa 1

03/11/2026. Definiciones preparadas, no desplegadas. Fuente autoritativa: archivos SQL individuales.

## Moneda

Fuente: `Tablas/Moneda.sql`

```sql
Codigo NVARCHAR(3) NOT NULL CONSTRAINT PK_Moneda PRIMARY KEY,
        Nombre NVARCHAR(40) NOT NULL,
        Simbolo NVARCHAR(5) NOT NULL
```

## TipoDocumento

Fuente: `Tablas/TipoDocumento.sql`

```sql
Codigo NVARCHAR(2) NOT NULL,
        Nombre NVARCHAR(40) NOT NULL,
        CONSTRAINT PK_TipoDocumento PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoDocumento_Codigo CHECK (Codigo IN (N'01', N'03'))
```

## TipoIGV

Fuente: `Tablas/TipoIGV.sql`

```sql
Codigo NVARCHAR(10) NOT NULL,
        Nombre NVARCHAR(20) NOT NULL,
        IncluyeIGV BIT NOT NULL,
        CONSTRAINT PK_TipoIGV PRIMARY KEY (Codigo),
        CONSTRAINT CK_TipoIGV_Tratamiento CHECK
        (
            (Codigo = N'CON_IGV' AND IncluyeIGV = 1)
            OR (Codigo = N'SIN_IGV' AND IncluyeIGV = 0)
        )
```

## TasaIGV

Fuente: `Tablas/TasaIGV.sql`

```sql
FechaDesde DATE NOT NULL,
        FechaHasta DATE NULL,
        Valor DECIMAL(9,6) NOT NULL,
        CONSTRAINT PK_TasaIGV PRIMARY KEY (FechaDesde),
        CONSTRAINT UQ_TasaIGV_Valor UNIQUE (FechaDesde, Valor),
        CONSTRAINT CK_TasaIGV_Valor CHECK (Valor >= 0 AND Valor <= 1),
        CONSTRAINT CK_TasaIGV_Vigencia CHECK (FechaHasta IS NULL OR FechaHasta > FechaDesde)
```

## AspNetUsers

Fuente: `Tablas/AspNetUsers.sql`

```sql
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
```

## AspNetRoles

Fuente: `Tablas/AspNetRoles.sql`

```sql
Id NVARCHAR(128) NOT NULL CONSTRAINT PK_AspNetRoles PRIMARY KEY,
        Name NVARCHAR(256) NULL,
        NormalizedName NVARCHAR(256) NULL,
        ConcurrencyStamp NVARCHAR(MAX) NULL
```

## AspNetUserRoles

Fuente: `Tablas/AspNetUserRoles.sql`

```sql
UserId NVARCHAR(128) NOT NULL,
        RoleId NVARCHAR(128) NOT NULL,
        CONSTRAINT PK_AspNetUserRoles PRIMARY KEY (UserId, RoleId),
        CONSTRAINT FK_UserRoles_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_UserRoles_Role FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id)
```

## AspNetUserClaims

Fuente: `Tablas/AspNetUserClaims.sql`

```sql
Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetUserClaims PRIMARY KEY,
        UserId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetUserClaims_Owner FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
```

## AspNetRoleClaims

Fuente: `Tablas/AspNetRoleClaims.sql`

```sql
Id INT IDENTITY NOT NULL CONSTRAINT PK_AspNetRoleClaims PRIMARY KEY,
        RoleId NVARCHAR(128) NOT NULL,
        ClaimType NVARCHAR(MAX) NULL,
        ClaimValue NVARCHAR(MAX) NULL,
        CONSTRAINT FK_AspNetRoleClaims_Owner FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id)
```

## AspNetUserLogins

Fuente: `Tablas/AspNetUserLogins.sql`

```sql
LoginProvider NVARCHAR(128) NOT NULL,
        ProviderKey NVARCHAR(128) NOT NULL,
        ProviderDisplayName NVARCHAR(MAX) NULL,
        UserId NVARCHAR(128) NOT NULL,
        CONSTRAINT PK_AspNetUserLogins PRIMARY KEY (LoginProvider, ProviderKey),
        CONSTRAINT FK_UserLogins_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
```

## AspNetUserTokens

Fuente: `Tablas/AspNetUserTokens.sql`

```sql
UserId NVARCHAR(128) NOT NULL,
        LoginProvider NVARCHAR(128) NOT NULL,
        Name NVARCHAR(128) NOT NULL,
        Value NVARCHAR(MAX) NULL,
        CONSTRAINT PK_AspNetUserTokens PRIMARY KEY (UserId, LoginProvider, Name),
        CONSTRAINT FK_UserTokens_User FOREIGN KEY (UserId) REFERENCES dbo.AspNetUsers(Id)
```

## Permiso

Fuente: `Tablas/Permiso.sql`

```sql
Codigo NVARCHAR(80) NOT NULL CONSTRAINT PK_Permiso PRIMARY KEY,
        Descripcion NVARCHAR(200) NOT NULL
```

## RolPermiso

Fuente: `Tablas/RolPermiso.sql`

```sql
RoleId NVARCHAR(128) NOT NULL,
        CodigoPermiso NVARCHAR(80) NOT NULL,
        CONSTRAINT PK_RolPermiso PRIMARY KEY (RoleId, CodigoPermiso),
        CONSTRAINT FK_RolPermiso_Rol FOREIGN KEY (RoleId) REFERENCES dbo.AspNetRoles(Id),
        CONSTRAINT FK_RolPermiso_Permiso FOREIGN KEY (CodigoPermiso) REFERENCES dbo.Permiso(Codigo)
```

## EventoAuditoria

Fuente: `Tablas/EventoAuditoria.sql`

```sql
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
```

## ConfiguracionHYT

Fuente: `Tablas/ConfiguracionHYT.sql`

```sql
Id TINYINT NOT NULL CONSTRAINT PK_ConfiguracionHYT PRIMARY KEY,
        Ruc NVARCHAR(20) NOT NULL,
        RazonSocial NVARCHAR(200) NOT NULL,
        Direccion NVARCHAR(300) NULL,
        Telefono NVARCHAR(50) NULL,
        Correo NVARCHAR(256) NULL,
        NombreFirmante NVARCHAR(150) NULL,
        CargoFirmante NVARCHAR(150) NULL,
        CONSTRAINT CK_ConfiguracionHYT_Unica CHECK (Id = 1),
        FechaRegistro DATETIME2(3) NOT NULL CONSTRAINT DF_ConfiguracionHYT_Registro DEFAULT SYSUTCDATETIME(),
        UsuarioRegistro NVARCHAR(128) NOT NULL,
        FechaActualizacion DATETIME2(3) NULL,
        UsuarioActualizacion NVARCHAR(128) NULL,
        VersionFila ROWVERSION NOT NULL,
        CONSTRAINT FK_ConfiguracionHYT_Autor FOREIGN KEY (UsuarioRegistro) REFERENCES dbo.AspNetUsers(Id),
        CONSTRAINT FK_ConfiguracionHYT_Editor FOREIGN KEY (UsuarioActualizacion) REFERENCES dbo.AspNetUsers(Id)
```

## UnidadMedida

Fuente: `Tablas/UnidadMedida.sql`

```sql
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
```

## Zona

Fuente: `Tablas/Zona.sql`

```sql
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
```

## Especialidad

Fuente: `Tablas/Especialidad.sql`

```sql
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
```

## TipoServicio

Fuente: `Tablas/TipoServicio.sql`

```sql
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
```

## Cliente

Fuente: `Tablas/Cliente.sql`

```sql
ClienteId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Cliente PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Cliente_Codigo UNIQUE,
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
```

## Tienda

Fuente: `Tablas/Tienda.sql`

```sql
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
```

## ContactoCliente

Fuente: `Tablas/ContactoCliente.sql`

```sql
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
```

## ProveedorCliente

Fuente: `Tablas/ProveedorCliente.sql`

```sql
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
```

## ServicioMaterial

Fuente: `Tablas/ServicioMaterial.sql`

```sql
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
```

## MaterialZona

Fuente: `Tablas/MaterialZona.sql`

```sql
MaterialId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_MaterialZona PRIMARY KEY (MaterialId, ZonaId),
        CONSTRAINT FK_MaterialZona_Material FOREIGN KEY (MaterialId) REFERENCES dbo.ServicioMaterial(MaterialId),
        CONSTRAINT FK_MaterialZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
```

## Tarifario

Fuente: `Tablas/Tarifario.sql`

```sql
TarifarioId BIGINT IDENTITY NOT NULL CONSTRAINT PK_Tarifario PRIMARY KEY,
        Codigo NVARCHAR(40) NOT NULL CONSTRAINT UQ_Tarifario_Codigo UNIQUE,
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
```

## TarifarioVersion

Fuente: `Tablas/TarifarioVersion.sql`

```sql
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
```

## TarifarioVersionZona

Fuente: `Tablas/TarifarioVersionZona.sql`

```sql
VersionId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioVersionZona PRIMARY KEY (VersionId, ZonaId),
        CONSTRAINT FK_TarifarioVersionZona_Version FOREIGN KEY (VersionId) REFERENCES dbo.TarifarioVersion(VersionId),
        CONSTRAINT FK_TarifarioVersionZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
```

## TarifarioPartida

Fuente: `Tablas/TarifarioPartida.sql`

```sql
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
```

## TarifarioPartidaZona

Fuente: `Tablas/TarifarioPartidaZona.sql`

```sql
PartidaId BIGINT NOT NULL,
        ZonaId INT NOT NULL,
        CONSTRAINT PK_TarifarioPartidaZona PRIMARY KEY (PartidaId, ZonaId),
        CONSTRAINT FK_TarifarioPartidaZona_Partida FOREIGN KEY (PartidaId) REFERENCES dbo.TarifarioPartida(PartidaId),
        CONSTRAINT FK_TarifarioPartidaZona_Zona FOREIGN KEY (ZonaId) REFERENCES dbo.Zona(Id)
```

## Presupuesto

Fuente: `Tablas/Presupuesto.sql`

```sql
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
```

## PresupuestoRevision

Fuente: `Tablas/PresupuestoRevision.sql`

```sql
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
```

## PresupuestoPartida

Fuente: `Tablas/PresupuestoPartida.sql`

```sql
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
        TipoServicio NVARCHAR(60) NULL,
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
```

## OrdenCompraCliente

Fuente: `Tablas/OrdenCompraCliente.sql`

```sql
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
```

## Ejecucion

Fuente: `Tablas/Ejecucion.sql`

```sql
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
```

## GuiaConformidad

Fuente: `Tablas/GuiaConformidad.sql`

```sql
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
```

## GuiaPartida

Fuente: `Tablas/GuiaPartida.sql`

```sql
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
```

## SeguimientoCliente

Fuente: `Tablas/SeguimientoCliente.sql`

```sql
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
```

## ConfirmacionFacturacion

Fuente: `Tablas/ConfirmacionFacturacion.sql`

```sql
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
```

## ReferenciaComprobanteExterno

Fuente: `Tablas/ReferenciaComprobanteExterno.sql`

```sql
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
```

## ArchivoAdjunto

Fuente: `Tablas/ArchivoAdjunto.sql`

```sql
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
```

## OrdenCompraArchivo

Fuente: `Tablas/OrdenCompraArchivo.sql`

```sql
OrdenCompraId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_OrdenCompraArchivo PRIMARY KEY (OrdenCompraId, ArchivoId),
        CONSTRAINT FK_OrdenCompraArchivo_Owner FOREIGN KEY (OrdenCompraId) REFERENCES dbo.OrdenCompraCliente(OrdenCompraId),
        CONSTRAINT FK_OrdenCompraArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
```

## GuiaArchivo

Fuente: `Tablas/GuiaArchivo.sql`

```sql
GuiaId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_GuiaArchivo PRIMARY KEY (GuiaId, ArchivoId),
        CONSTRAINT FK_GuiaArchivo_Owner FOREIGN KEY (GuiaId) REFERENCES dbo.GuiaConformidad(GuiaId),
        CONSTRAINT FK_GuiaArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
```

## ConfirmacionArchivo

Fuente: `Tablas/ConfirmacionArchivo.sql`

```sql
ConfirmacionId BIGINT NOT NULL,
        ArchivoId BIGINT NOT NULL,
        CONSTRAINT PK_ConfirmacionArchivo PRIMARY KEY (ConfirmacionId, ArchivoId),
        CONSTRAINT FK_ConfirmacionArchivo_Owner FOREIGN KEY (ConfirmacionId) REFERENCES dbo.ConfirmacionFacturacion(ConfirmacionId),
        CONSTRAINT FK_ConfirmacionArchivo_Archivo FOREIGN KEY (ArchivoId) REFERENCES dbo.ArchivoAdjunto(ArchivoId)
```

## ImpresionPresupuesto

Fuente: `Tablas/ImpresionPresupuesto.sql`

```sql
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
```

## NumeracionInterna

Fuente: `Tablas/NumeracionInterna.sql`

```sql
TipoRegistro NVARCHAR(40) NOT NULL CONSTRAINT PK_NumeracionInterna PRIMARY KEY,
        Prefijo NVARCHAR(20) NOT NULL,
        UltimoNumero BIGINT NOT NULL DEFAULT 0,
        CONSTRAINT CK_Numeracion_Numero CHECK (UltimoNumero >= 0)
```

