-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Registra aplicaciones parciales o totales entre comprobantes pendientes y notas de credito de compras o ventas.
-- =============================================
-- Firma: FRANCO LARA - 24/06/2026 | Crea la tabla del modulo Aplicaciones para enlazar un comprobante con una nota de credito, guardar el importe aplicado y vincular el asiento generado.

IF OBJECT_ID(N'dbo.CON_AplicacionNotaCredito', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CON_AplicacionNotaCredito
    (
        IdAplicacionNotaCredito INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CON_AplicacionNotaCredito PRIMARY KEY,
        IdEmpresa INT NOT NULL,
        ModuloOperacion VARCHAR(10) NOT NULL,
        IdPersona INT NOT NULL,
        FechaAplicacion DATE NOT NULL,
        IdRegistroComprobante INT NOT NULL,
        IdRegistroNotaCredito INT NOT NULL,
        IdMoneda INT NOT NULL,
        TipoCambio DECIMAL(18, 6) NOT NULL CONSTRAINT DF_CON_AplicacionNotaCredito_TipoCambio DEFAULT (1),
        ImporteAplicado DECIMAL(18, 2) NOT NULL,
        IdAsiento INT NULL,
        Glosa NVARCHAR(300) NOT NULL,
        Observacion NVARCHAR(500) NULL,
        Activo BIT NOT NULL CONSTRAINT DF_CON_AplicacionNotaCredito_Activo DEFAULT (1),
        FechaRegistro DATETIME2(0) NOT NULL CONSTRAINT DF_CON_AplicacionNotaCredito_FechaRegistro DEFAULT (SYSDATETIME()),
        UsuarioRegistro NVARCHAR(450) NULL
    );

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT FK_CON_AplicacionNotaCredito_SEG_Empresa
            FOREIGN KEY (IdEmpresa) REFERENCES dbo.SEG_Empresa (IdEmpresa);

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT FK_CON_AplicacionNotaCredito_ADM_Persona
            FOREIGN KEY (IdPersona) REFERENCES dbo.ADM_Persona (IdPersona);

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT FK_CON_AplicacionNotaCredito_ADM_Moneda
            FOREIGN KEY (IdMoneda) REFERENCES dbo.ADM_Moneda (IdMoneda);

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT FK_CON_AplicacionNotaCredito_CON_Asiento
            FOREIGN KEY (IdAsiento) REFERENCES dbo.CON_Asiento (IdAsiento);

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT CK_CON_AplicacionNotaCredito_ModuloOperacion
            CHECK (ModuloOperacion IN ('COM', 'VEN'));

    ALTER TABLE dbo.CON_AplicacionNotaCredito
        ADD CONSTRAINT CK_CON_AplicacionNotaCredito_ImporteAplicado
            CHECK (ImporteAplicado > 0);
END;
