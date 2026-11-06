-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Agrega vigencia futura de IGV y cierra la anterior con auditoria.
-- =============================================
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
