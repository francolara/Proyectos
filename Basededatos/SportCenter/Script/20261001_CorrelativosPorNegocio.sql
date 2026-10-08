
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/10/2026
-- Description:   Agrega la estructura para correlativos visibles e independientes por negocio.
--                No genera correlativos para registros historicos y no utiliza triggers.
-- =============================================
SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRANSACTION;

    IF OBJECT_ID(N'dbo.NegocioCorrelativos', N'U') IS NULL
    BEGIN
        CREATE TABLE dbo.NegocioCorrelativos
        (
            NegocioId INT NOT NULL,
            Entidad NVARCHAR(30) NOT NULL,
            UltimoNumero INT NOT NULL,
            FechaActualizacion DATETIME2(7) NOT NULL
                CONSTRAINT DF_NegocioCorrelativos_FechaActualizacion DEFAULT (SYSUTCDATETIME()),
            UsuarioActualizacion NVARCHAR(200) NULL,
            CONSTRAINT PK_NegocioCorrelativos
                PRIMARY KEY CLUSTERED (NegocioId ASC, Entidad ASC),
            CONSTRAINT CK_NegocioCorrelativos_Entidad
                CHECK (Entidad IN (N'RESERVA', N'PAGO', N'CLIENTE', N'SEDE', N'ESPACIO', N'CUPON', N'PROMOCION')),
            CONSTRAINT CK_NegocioCorrelativos_UltimoNumero
                CHECK (UltimoNumero >= 0),
            CONSTRAINT FK_NegocioCorrelativos_Negocios_NegocioId
                FOREIGN KEY (NegocioId) REFERENCES dbo.Negocios (Id) ON DELETE CASCADE
        );
    END;

    IF COL_LENGTH(N'dbo.Reservas', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Reservas ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Pagos', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Pagos ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Clientes', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Clientes ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Sedes', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Sedes ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.EspaciosDeportivos', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.EspaciosDeportivos ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.Cupones', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.Cupones ADD NumeroPorNegocio INT NULL;

    IF COL_LENGTH(N'dbo.PromocionesHorario', N'NumeroPorNegocio') IS NULL
        ALTER TABLE dbo.PromocionesHorario ADD NumeroPorNegocio INT NULL;

    COMMIT TRANSACTION;
END TRY
BEGIN CATCH
    IF XACT_STATE() <> 0
        ROLLBACK TRANSACTION;

    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
GO
