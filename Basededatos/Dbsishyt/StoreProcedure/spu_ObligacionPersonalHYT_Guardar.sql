-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ObligacionPersonalHYT_Guardar para obligaciones de Planillas.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Exige alta pendiente y conserva pagos anteriores al editar la programacion.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ObligacionPersonalHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Tipo VARCHAR(3),
    @Entidad NVARCHAR(120),
    @Mes DATE,
    @Importe DECIMAL(20,4),
    @Vencimiento DATE,
    @FechaPago DATE = NULL,
    @ReferenciaPago NVARCHAR(100) = NULL,
    @ObligacionId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ConservarPago BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar obligaciones.', 1;

        IF @ConservarPago = 1
        BEGIN
            SET @FechaPago = NULL;
            SET @ReferenciaPago = NULL;
        END;
        SET @Entidad = NULLIF(TRIM(@Entidad), N'');
        SET @ReferenciaPago = NULLIF(TRIM(@ReferenciaPago), N'');
        DECLARE @Hoy DATE = CONVERT(DATE, DATEADD(HOUR, -5, SYSUTCDATETIME()));
        IF @Tipo IS NULL OR @Tipo NOT IN ('AFP', 'ONP') OR @Entidad IS NULL OR @Mes IS NULL
           OR DAY(@Mes) <> 1 OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
           OR @Vencimiento IS NULL
           OR (@ObligacionId IS NULL AND @FechaPago IS NOT NULL)
           OR (@FechaPago IS NULL AND @ReferenciaPago IS NOT NULL)
           OR (@FechaPago IS NOT NULL AND (@ReferenciaPago IS NULL OR @FechaPago > @Hoy))
            THROW 51000, N'Revise la entidad, el mes, el importe y las fechas de la obligacion.', 1;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @ObligacionId IS NULL
        BEGIN
            INSERT INTO dbo.ObligacionPersonalHYT (Tipo, Entidad, Mes, Importe, Vencimiento, FechaPago, ReferenciaPago)
            VALUES (@Tipo, @Entidad, @Mes, @Importe, @Vencimiento, @FechaPago, @ReferenciaPago);
            SET @ObligacionId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            IF @ConservarPago = 1
                SELECT @FechaPago = obligacion.FechaPago, @ReferenciaPago = obligacion.ReferenciaPago
                FROM dbo.ObligacionPersonalHYT AS obligacion WITH (UPDLOCK)
                WHERE obligacion.ObligacionId = @ObligacionId AND obligacion.VersionFila = @VersionFila;

            UPDATE dbo.ObligacionPersonalHYT
            SET Tipo = @Tipo, Entidad = @Entidad, Mes = @Mes, Importe = @Importe, Vencimiento = @Vencimiento,
                FechaPago = @FechaPago, ReferenciaPago = @ReferenciaPago
            WHERE ObligacionId = @ObligacionId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La obligacion cambio o no existe. Vuelva a abrirla.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @ObligacionId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Tipo AS Tipo, @Entidad AS Entidad, @Mes AS Mes, @Importe AS Importe,
                   @Vencimiento AS Vencimiento, @FechaPago AS FechaPago, @ReferenciaPago AS ReferenciaPago
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER, INCLUDE_NULL_VALUES
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'ObligacionPersonalHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @ObligacionId AS ObligacionId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
