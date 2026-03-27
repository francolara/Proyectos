-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Creacion de tablas de modulos y permisos por rol/usuario para administracion por negocio.
-- =============================================
BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE TABLE [ModulosSistema] (
        [Id] int NOT NULL IDENTITY,
        [Codigo] nvarchar(50) NOT NULL,
        [Nombre] nvarchar(120) NOT NULL,
        [Activo] bit NOT NULL,
        CONSTRAINT [PK_ModulosSistema] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE TABLE [RolesNegocioPermiso] (
        [Id] int NOT NULL IDENTITY,
        [RolNegocio] int NOT NULL,
        [ModuloSistemaId] int NOT NULL,
        [PuedeVer] bit NOT NULL,
        [PuedeCrear] bit NOT NULL,
        [PuedeEditar] bit NOT NULL,
        [PuedeEliminar] bit NOT NULL,
        CONSTRAINT [PK_RolesNegocioPermiso] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_RolesNegocioPermiso_ModulosSistema_ModuloSistemaId] FOREIGN KEY ([ModuloSistemaId]) REFERENCES [ModulosSistema] ([Id]) ON DELETE NO ACTION
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE TABLE [UsuariosNegocioPermiso] (
        [Id] int NOT NULL IDENTITY,
        [UsuarioNegocioId] int NOT NULL,
        [ModuloSistemaId] int NOT NULL,
        [PuedeVer] bit NOT NULL,
        [PuedeCrear] bit NOT NULL,
        [PuedeEditar] bit NOT NULL,
        [PuedeEliminar] bit NOT NULL,
        CONSTRAINT [PK_UsuariosNegocioPermiso] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_UsuariosNegocioPermiso_ModulosSistema_ModuloSistemaId] FOREIGN KEY ([ModuloSistemaId]) REFERENCES [ModulosSistema] ([Id]) ON DELETE NO ACTION,
        CONSTRAINT [FK_UsuariosNegocioPermiso_UsuariosNegocio_UsuarioNegocioId] FOREIGN KEY ([UsuarioNegocioId]) REFERENCES [UsuariosNegocio] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE UNIQUE INDEX [IX_ModulosSistema_Codigo] ON [ModulosSistema] ([Codigo]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE INDEX [IX_RolesNegocioPermiso_ModuloSistemaId] ON [RolesNegocioPermiso] ([ModuloSistemaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE UNIQUE INDEX [IX_RolesNegocioPermiso_RolNegocio_ModuloSistemaId] ON [RolesNegocioPermiso] ([RolNegocio], [ModuloSistemaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE INDEX [IX_UsuariosNegocioPermiso_ModuloSistemaId] ON [UsuariosNegocioPermiso] ([ModuloSistemaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    CREATE UNIQUE INDEX [IX_UsuariosNegocioPermiso_UsuarioNegocioId_ModuloSistemaId] ON [UsuariosNegocioPermiso] ([UsuarioNegocioId], [ModuloSistemaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326224945_AddNegocioPermissions'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260326224945_AddNegocioPermissions', N'10.0.5');
END;

COMMIT;
GO

