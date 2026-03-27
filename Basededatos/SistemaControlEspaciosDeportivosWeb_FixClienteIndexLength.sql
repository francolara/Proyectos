-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Ajuste de longitud en columnas de Clientes para evitar advertencia por tamano maximo de indice.
-- =============================================
BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    DROP INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON [Clientes];
    DECLARE @var nvarchar(max);
    SELECT @var = QUOTENAME([d].[name])
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Clientes]') AND [c].[name] = N'TipoDocumento');
    IF @var IS NOT NULL EXEC(N'ALTER TABLE [Clientes] DROP CONSTRAINT ' + @var + ';');
    ALTER TABLE [Clientes] ALTER COLUMN [TipoDocumento] nvarchar(20) NOT NULL;
    CREATE UNIQUE INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON [Clientes] ([TipoDocumento], [NumeroDocumento]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    DECLARE @var1 nvarchar(max);
    SELECT @var1 = QUOTENAME([d].[name])
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Clientes]') AND [c].[name] = N'Telefono');
    IF @var1 IS NOT NULL EXEC(N'ALTER TABLE [Clientes] DROP CONSTRAINT ' + @var1 + ';');
    ALTER TABLE [Clientes] ALTER COLUMN [Telefono] nvarchar(20) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    DROP INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON [Clientes];
    DECLARE @var2 nvarchar(max);
    SELECT @var2 = QUOTENAME([d].[name])
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Clientes]') AND [c].[name] = N'NumeroDocumento');
    IF @var2 IS NOT NULL EXEC(N'ALTER TABLE [Clientes] DROP CONSTRAINT ' + @var2 + ';');
    ALTER TABLE [Clientes] ALTER COLUMN [NumeroDocumento] nvarchar(20) NOT NULL;
    CREATE UNIQUE INDEX [IX_Clientes_TipoDocumento_NumeroDocumento] ON [Clientes] ([TipoDocumento], [NumeroDocumento]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    DECLARE @var3 nvarchar(max);
    SELECT @var3 = QUOTENAME([d].[name])
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Clientes]') AND [c].[name] = N'NombresORazonSocial');
    IF @var3 IS NOT NULL EXEC(N'ALTER TABLE [Clientes] DROP CONSTRAINT ' + @var3 + ';');
    ALTER TABLE [Clientes] ALTER COLUMN [NombresORazonSocial] nvarchar(200) NOT NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    DECLARE @var4 nvarchar(max);
    SELECT @var4 = QUOTENAME([d].[name])
    FROM [sys].[default_constraints] [d]
    INNER JOIN [sys].[columns] [c] ON [d].[parent_column_id] = [c].[column_id] AND [d].[parent_object_id] = [c].[object_id]
    WHERE ([d].[parent_object_id] = OBJECT_ID(N'[Clientes]') AND [c].[name] = N'Correo');
    IF @var4 IS NOT NULL EXEC(N'ALTER TABLE [Clientes] DROP CONSTRAINT ' + @var4 + ';');
    ALTER TABLE [Clientes] ALTER COLUMN [Correo] nvarchar(200) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326222934_FixClienteIndexLength'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260326222934_FixClienteIndexLength', N'10.0.5');
END;

COMMIT;
GO

