-- =============================================
-- Author:        FRANCO LARA
-- Create date:   26/03/2026
-- Description:   Creacion de tablas y relaciones para boleta/factura electronica asociada a reservas.
-- =============================================
BEGIN TRANSACTION;
IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    ALTER TABLE [Clientes] ADD [DireccionFiscal] nvarchar(250) NULL;
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE TABLE [ComprobantesElectronicos] (
        [Id] int NOT NULL IDENTITY,
        [NegocioId] int NOT NULL,
        [ReservaId] int NOT NULL,
        [ClienteId] int NOT NULL,
        [TipoComprobante] int NOT NULL,
        [Serie] nvarchar(4) NOT NULL,
        [Numero] int NOT NULL,
        [FechaEmision] datetime2 NOT NULL,
        [TipoMoneda] int NOT NULL,
        [CodigoTipoOperacionSunat] nvarchar(4) NOT NULL,
        [CodigoTipoDocumentoClienteSunat] nvarchar(4) NOT NULL,
        [CodigoHashCpe] nvarchar(8) NULL,
        [NumeroTicketSunat] nvarchar(40) NULL,
        [CodigoRespuestaSunat] nvarchar(50) NULL,
        [MensajeRespuestaSunat] nvarchar(500) NULL,
        [SubTotal] decimal(10,2) NOT NULL,
        [Igv] decimal(10,2) NOT NULL,
        [Total] decimal(10,2) NOT NULL,
        [Estado] int NOT NULL,
        [FechaRegistro] datetime2 NOT NULL,
        CONSTRAINT [PK_ComprobantesElectronicos] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_ComprobantesElectronicos_Clientes_ClienteId] FOREIGN KEY ([ClienteId]) REFERENCES [Clientes] ([Id]) ON DELETE NO ACTION,
        CONSTRAINT [FK_ComprobantesElectronicos_Negocios_NegocioId] FOREIGN KEY ([NegocioId]) REFERENCES [Negocios] ([Id]) ON DELETE NO ACTION,
        CONSTRAINT [FK_ComprobantesElectronicos_Reservas_ReservaId] FOREIGN KEY ([ReservaId]) REFERENCES [Reservas] ([Id]) ON DELETE NO ACTION
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE TABLE [ComprobantesDetalle] (
        [Id] int NOT NULL IDENTITY,
        [ComprobanteElectronicoId] int NOT NULL,
        [Item] int NOT NULL,
        [Descripcion] nvarchar(250) NOT NULL,
        [Cantidad] decimal(10,2) NOT NULL,
        [UnidadMedidaSunat] nvarchar(3) NOT NULL,
        [ValorUnitario] decimal(10,2) NOT NULL,
        [PrecioUnitario] decimal(10,2) NOT NULL,
        [BaseIgv] decimal(10,2) NOT NULL,
        [Igv] decimal(10,2) NOT NULL,
        [Total] decimal(10,2) NOT NULL,
        [AfectacionIgvSunat] nvarchar(2) NOT NULL,
        CONSTRAINT [PK_ComprobantesDetalle] PRIMARY KEY ([Id]),
        CONSTRAINT [FK_ComprobantesDetalle_ComprobantesElectronicos_ComprobanteElectronicoId] FOREIGN KEY ([ComprobanteElectronicoId]) REFERENCES [ComprobantesElectronicos] ([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE UNIQUE INDEX [IX_ComprobantesDetalle_ComprobanteElectronicoId_Item] ON [ComprobantesDetalle] ([ComprobanteElectronicoId], [Item]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE INDEX [IX_ComprobantesElectronicos_ClienteId] ON [ComprobantesElectronicos] ([ClienteId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE UNIQUE INDEX [IX_ComprobantesElectronicos_NegocioId_TipoComprobante_Serie_Numero] ON [ComprobantesElectronicos] ([NegocioId], [TipoComprobante], [Serie], [Numero]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    CREATE UNIQUE INDEX [IX_ComprobantesElectronicos_ReservaId] ON [ComprobantesElectronicos] ([ReservaId]);
END;

IF NOT EXISTS (
    SELECT * FROM [__EFMigrationsHistory]
    WHERE [MigrationId] = N'20260326223450_AddComprobanteElectronico'
)
BEGIN
    INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260326223450_AddComprobanteElectronico', N'10.0.5');
END;

COMMIT;
GO

