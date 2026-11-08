-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Finaliza intento con bloqueo persistente y auditoria atomica.
-- =============================================
-- Firma: FRANCO LARA | Create date: 06/11/2026 | Último ingreso y bloqueo automático auditado.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_FinalizarAcceso
    @UsuarioId NVARCHAR(128),
    @SelloEsperado NVARCHAR(MAX),
    @PasswordValido BIT,
    @HashActualizado NVARCHAR(MAX) = NULL,
    @CorrelacionId UNIQUEIDENTIFIER
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF @@TRANCOUNT <> 0 THROW 51101, N'El acceso administra su propia transaccion.', 1;
        DECLARE @Activo BIT, @Sello NVARCHAR(MAX), @Bloqueo DATETIMEOFFSET,
            @Intentos INT, @Exito BIT = 0, @Ahora DATETIMEOFFSET = SYSUTCDATETIME();
        BEGIN TRANSACTION;
        SET @Propia = 1;
        SELECT @Activo = usuario.Activo, @Sello = COALESCE(usuario.SecurityStamp, N''),
            @Bloqueo = usuario.LockoutEnd, @Intentos = usuario.AccessFailedCount
        FROM dbo.AspNetUsers AS usuario WITH (UPDLOCK, HOLDLOCK) WHERE usuario.Id = @UsuarioId;
        IF @Activo = 1 AND @Sello = @SelloEsperado AND (@Bloqueo IS NULL OR @Bloqueo <= @Ahora)
        BEGIN
            IF @PasswordValido = 1
            BEGIN
                UPDATE dbo.AspNetUsers SET AccessFailedCount = 0, LockoutEnd = NULL,
                    PasswordHash = COALESCE(@HashActualizado, PasswordHash),UltimoIngreso=SYSUTCDATETIME()
                WHERE Id = @UsuarioId;
                SET @Exito = 1;
            END
            ELSE
            BEGIN
                SET @Intentos = CASE WHEN @Bloqueo IS NOT NULL THEN 1 ELSE @Intentos + 1 END;
                UPDATE dbo.AspNetUsers SET AccessFailedCount = CASE WHEN @Intentos >= 5 THEN 0 ELSE @Intentos END,
                    LockoutEnd = CASE WHEN @Intentos >= 5 THEN DATEADD(MINUTE, 15, @Ahora) ELSE NULL END
                WHERE Id = @UsuarioId;
            END;
        END;
        IF @Activo IS NOT NULL
        BEGIN
            INSERT INTO dbo.EventoAuditoria (CorrelacionId, Entidad, RegistroId, Accion, UsuarioId)
            VALUES (@CorrelacionId, N'AspNetUsers', @UsuarioId,
                CASE WHEN @Exito = 1 THEN N'INICIO_SESION' ELSE N'ACCESO_RECHAZADO' END, @UsuarioId);
        END;
        IF @PasswordValido=0 AND @Intentos>=5 AND @Activo=1
            INSERT INTO dbo.EventoAuditoria(CorrelacionId,Entidad,RegistroId,Accion,UsuarioId) VALUES(@CorrelacionId,N'AspNetUsers',@UsuarioId,N'BLOQUEO_AUTOMATICO',@UsuarioId);
        COMMIT TRANSACTION;
        SELECT @Exito AS Exito;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
