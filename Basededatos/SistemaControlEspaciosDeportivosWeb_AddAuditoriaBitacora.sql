-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Creacion de bitacora de auditoria y campos de trazabilidad en entidades operativas.
-- =============================================
BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Sedes] ADD [FechaActualizacion] datetime2 NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Sedes] ADD [FechaCreacion] datetime2 NOT NULL DEFAULT '0001-01-01T00:00:00.0000000';
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Sedes] ADD [UsuarioActualizacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Sedes] ADD [UsuarioCreacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Reservas] ADD [FechaActualizacion] datetime2 NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Reservas] ADD [UsuarioActualizacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Reservas] ADD [UsuarioCreacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Pagos] ADD [FechaActualizacion] datetime2 NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Pagos] ADD [FechaCreacion] datetime2 NOT NULL DEFAULT '0001-01-01T00:00:00.0000000';
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Pagos] ADD [UsuarioActualizacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [Pagos] ADD [UsuarioCreacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [EspaciosDeportivos] ADD [FechaActualizacion] datetime2 NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [EspaciosDeportivos] ADD [FechaCreacion] datetime2 NOT NULL DEFAULT '0001-01-01T00:00:00.0000000';
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [EspaciosDeportivos] ADD [UsuarioActualizacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [EspaciosDeportivos] ADD [UsuarioCreacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [ComprobantesElectronicos] ADD [FechaActualizacion] datetime2 NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [ComprobantesElectronicos] ADD [UsuarioActualizacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    ALTER TABLE [ComprobantesElectronicos] ADD [UsuarioCreacion] nvarchar(max) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    CREATE TABLE [BitacoraAuditoria] (
        [Id] bigint NOT NULL IDENTITY,
        [NegocioId] int NULL,
        [Modulo] nvarchar(50) NOT NULL,
        [Accion] nvarchar(20) NOT NULL,
        [Entidad] nvarchar(80) NOT NULL,
        [EntidadId] nvarchar(80) NOT NULL,
        [UsuarioId] nvarchar(450) NOT NULL,
        [UsuarioNombre] nvarchar(200) NULL,
        [UsuarioCorreo] nvarchar(200) NULL,
        [DetalleJson] nvarchar(4000) NULL,
        [FechaRegistro] datetime2 NOT NULL,
        CONSTRAINT [PK_BitacoraAuditoria] PRIMARY KEY ([Id])
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    CREATE INDEX [IX_BitacoraAuditoria_NegocioId_Modulo_FechaRegistro] ON [BitacoraAuditoria] ([NegocioId], [Modulo], [FechaRegistro]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260327010527_AddAuditoriaBitacora'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260327010527_AddAuditoriaBitacora', N'10.0.5');
END;

COMMIT;
GO

