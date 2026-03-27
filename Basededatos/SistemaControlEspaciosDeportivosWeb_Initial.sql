-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Script inicial idempotente de creacion de base y tablas para SistemaControlEspaciosDeportivosWeb.
-- =============================================
IF OBJECT_ID(N'[__EFMigrationsHistory]') IS NULL
BEGIN
    CREATE TABLE [__EFMigrationsHistory] (
        [MigrationId] nvarchar(150) NOT NULL,
        [ProductVersion] nvarchar(32) NOT NULL,
        CONSTRAINT [PK___EFMigrationsHistory] PRIMARY KEY ([MigrationId])
    );
END;
GO

BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetRoles] (
        [Id] nvarchar(450) NOT NULL,
        [Name] nvarchar(256) NULL,
        [NormalizedName] nvarchar(256) NULL,
        [ConcurrencyStamp] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetRoles] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetUsers] (
        [Id] nvarchar(450) NOT NULL,
        [UserName] nvarchar(256) NULL,
        [NormalizedUserName] nvarchar(256) NULL,
        [Email] nvarchar(256) NULL,
        [NormalizedEmail] nvarchar(256) NULL,
        [EmailConfirmed] bit NOT NULL,
        [PasswordHash] nvarchar(max) NULL,
        [SecurityStamp] nvarchar(max) NULL,
        [ConcurrencyStamp] nvarchar(max) NULL,
        [PhoneNumber] nvarchar(max) NULL,
        [PhoneNumberConfirmed] bit NOT NULL,
        [TwoFactorEnabled] bit NOT NULL,
        [LockoutEnd] datetimeoffset NULL,
        [LockoutEnabled] bit NOT NULL,
        [AccessFailedCount] int NOT NULL,
        CONSTRAINT [PK_AspNetUsers] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetRoleClaims] (
        [Id] int NOT NULL IDENTITY,
        [RoleId] nvarchar(450) NOT NULL,
        [ClaimType] nvarchar(max) NULL,
        [ClaimValue] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetRoleClaims] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_AspNetRoleClaims_AspNetRoles_RoleId] FOREIGN KEY ([RoleId]) REFERENCES [AspNetRoles] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetUserClaims] (
        [Id] int NOT NULL IDENTITY,
        [UserId] nvarchar(450) NOT NULL,
        [ClaimType] nvarchar(max) NULL,
        [ClaimValue] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetUserClaims] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_AspNetUserClaims_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetUserLogins] (
        [LoginProvider] nvarchar(128) NOT NULL,
        [ProviderKey] nvarchar(128) NOT NULL,
        [ProviderDisplayName] nvarchar(max) NULL,
        [UserId] nvarchar(450) NOT NULL,
        CONSTRAINT [PK_AspNetUserLogins] PRIMARY KEY ([LoginProvider], [ProviderKey]),
        CONSTRAINT [FK_AspNetUserLogins_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetUserRoles] (
        [UserId] nvarchar(450) NOT NULL,
        [RoleId] nvarchar(450) NOT NULL,
        CONSTRAINT [PK_AspNetUserRoles] PRIMARY KEY ([UserId], [RoleId]),
        CONSTRAINT [FK_AspNetUserRoles_AspNetRoles_RoleId] FOREIGN KEY ([RoleId]) REFERENCES [AspNetRoles] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_AspNetUserRoles_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE TABLE [AspNetUserTokens] (
        [UserId] nvarchar(450) NOT NULL,
        [LoginProvider] nvarchar(128) NOT NULL,
        [Name] nvarchar(128) NOT NULL,
        [Value] nvarchar(max) NULL,
        CONSTRAINT [PK_AspNetUserTokens] PRIMARY KEY ([UserId], [LoginProvider], [Name]),
        CONSTRAINT [FK_AspNetUserTokens_AspNetUsers_UserId] FOREIGN KEY ([UserId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE INDEX [IX_AspNetRoleClaims_RoleId] ON [AspNetRoleClaims] ([RoleId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [RoleNameIndex] ON [AspNetRoles] ([NormalizedName]) WHERE [NormalizedName] IS NOT NULL');
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE INDEX [IX_AspNetUserClaims_UserId] ON [AspNetUserClaims] ([UserId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE INDEX [IX_AspNetUserLogins_UserId] ON [AspNetUserLogins] ([UserId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE INDEX [IX_AspNetUserRoles_RoleId] ON [AspNetUserRoles] ([RoleId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    CREATE INDEX [EmailIndex] ON [AspNetUsers] ([NormalizedEmail]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    EXEC(N'CREATE UNIQUE INDEX [UserNameIndex] ON [AspNetUsers] ([NormalizedUserName]) WHERE [NormalizedUserName] IS NOT NULL');
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'00000000000000_CreateIdentitySchema'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'00000000000000_CreateIdentitySchema', N'10.0.5');
END;

COMMIT;
GO

BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    ALTER TABLE [AspNetUsers] ADD [Apellidos] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    ALTER TABLE [AspNetUsers] ADD [Nombres] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Clientes] (
        [Id] int NOT NULL IDENTITY,
        [NombresORazonSocial] nvarchar(max) NOT NULL,
        [TipoDocumento] nvarchar(450) NOT NULL,
        [NumeroDocumento] nvarchar(450) NOT NULL,
        [Telefono] nvarchar(max) NULL,
        [Correo] nvarchar(max) NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_Clientes] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Negocios] (
        [Id] int NOT NULL IDENTITY,
        [NombreComercial] nvarchar(max) NOT NULL,
        [RazonSocial] nvarchar(max) NULL,
        [DocumentoFiscal] nvarchar(max) NULL,
        [Activo] bit NOT NULL,
        [FechaRegistro] datetime2 NOT NULL,
        CONSTRAINT [PK_Negocios] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [TiposDeporte] (
        [Id] int NOT NULL IDENTITY,
        [Nombre] nvarchar(max) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_TiposDeporte] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Sedes] (
        [Id] int NOT NULL IDENTITY,
        [NegocioId] int NOT NULL,
        [Nombre] nvarchar(max) NOT NULL,
        [Direccion] nvarchar(max) NOT NULL,
        [Telefono] nvarchar(max) NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_Sedes] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Sedes_Negocios_NegocioId] FOREIGN KEY ([NegocioId]) REFERENCES [Negocios] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [UsuariosNegocio] (
        [Id] int NOT NULL IDENTITY,
        [UsuarioId] nvarchar(450) NOT NULL,
        [NegocioId] int NOT NULL,
        [RolNegocio] int NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_UsuariosNegocio] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_UsuariosNegocio_AspNetUsers_UsuarioId] FOREIGN KEY ([UsuarioId]) REFERENCES [AspNetUsers] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_UsuariosNegocio_Negocios_NegocioId] FOREIGN KEY ([NegocioId]) REFERENCES [Negocios] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [EspaciosDeportivos] (
        [Id] int NOT NULL IDENTITY,
        [SedeId] int NOT NULL,
        [TipoDeporteId] int NOT NULL,
        [Codigo] nvarchar(450) NOT NULL,
        [Nombre] nvarchar(max) NOT NULL,
        [Capacidad] int NOT NULL,
        [TieneIluminacion] bit NOT NULL,
        [Techada] bit NOT NULL,
        [Estado] int NOT NULL,
        CONSTRAINT [PK_EspaciosDeportivos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_EspaciosDeportivos_Sedes_SedeId] FOREIGN KEY ([SedeId]) REFERENCES [Sedes] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_EspaciosDeportivos_TiposDeporte_TipoDeporteId] FOREIGN KEY ([TipoDeporteId]) REFERENCES [TiposDeporte] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Reservas] (
        [Id] int NOT NULL IDENTITY,
        [EspacioDeportivoId] int NOT NULL,
        [ClienteId] int NOT NULL,
        [Fecha] date NOT NULL,
        [HoraInicio] time NOT NULL,
        [HoraFin] time NOT NULL,
        [Estado] int NOT NULL,
        [Total] decimal(10,2) NOT NULL,
        [Adelanto] decimal(10,2) NOT NULL,
        [Saldo] decimal(10,2) NOT NULL,
        [FechaRegistro] datetime2 NOT NULL,
        CONSTRAINT [PK_Reservas] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Reservas_Clientes_ClienteId] FOREIGN KEY ([ClienteId]) REFERENCES [Clientes] ([Id]) ON DELETE CASCADE,
        CONSTRAINT [FK_Reservas_EspaciosDeportivos_EspacioDeportivoId] FOREIGN KEY ([EspacioDeportivoId]) REFERENCES [EspaciosDeportivos] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Tarifas] (
        [Id] int NOT NULL IDENTITY,
        [EspacioDeportivoId] int NOT NULL,
        [DiaSemana] int NOT NULL,
        [HoraInicio] time NOT NULL,
        [HoraFin] time NOT NULL,
        [Precio] decimal(10,2) NOT NULL,
        [Activa] bit NOT NULL,
        CONSTRAINT [PK_Tarifas] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Tarifas_EspaciosDeportivos_EspacioDeportivoId] FOREIGN KEY ([EspacioDeportivoId]) REFERENCES [EspaciosDeportivos] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE TABLE [Pagos] (
        [Id] int NOT NULL IDENTITY,
        [ReservaId] int NOT NULL,
        [FechaPago] datetime2 NOT NULL,
        [Monto] decimal(10,2) NOT NULL,
        [FormaPago] int NOT NULL,
        [NumeroOperacion] nvarchar(max) NULL,
        [Observacion] nvarchar(max) NULL,
        CONSTRAINT [PK_Pagos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_Pagos_Reservas_ReservaId] FOREIGN KEY ([ReservaId]) REFERENCES [Reservas] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE UNIQUE INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON [Clientes] ([TipoDocumento], [NumeroDocumento]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE UNIQUE INDEX [IX_EspaciosDeportivos_SedeId_Codigo] ON [EspaciosDeportivos] ([SedeId], [Codigo]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_EspaciosDeportivos_TipoDeporteId] ON [EspaciosDeportivos] ([TipoDeporteId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_Pagos_ReservaId] ON [Pagos] ([ReservaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_Reservas_ClienteId] ON [Reservas] ([ClienteId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_Reservas_EspacioDeportivoId_Fecha_HoraInicio_HoraFin] ON [Reservas] ([EspacioDeportivoId], [Fecha], [HoraInicio], [HoraFin]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_Sedes_NegocioId] ON [Sedes] ([NegocioId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_Tarifas_EspacioDeportivoId_DiaSemana_HoraInicio_HoraFin] ON [Tarifas] ([EspacioDeportivoId], [DiaSemana], [HoraInicio], [HoraFin]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE INDEX [IX_UsuariosNegocio_NegocioId] ON [UsuariosNegocio] ([NegocioId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    CREATE UNIQUE INDEX [IX_UsuariosNegocio_UsuarioId_NegocioId] ON [UsuariosNegocio] ([UsuarioId], [NegocioId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326221335_InitialDomainStructure'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260326221335_InitialDomainStructure', N'10.0.5');
END;

COMMIT;
GO

