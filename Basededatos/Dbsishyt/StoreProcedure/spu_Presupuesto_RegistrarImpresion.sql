-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Conserva HTML, formato y huella de una revisión con control de concurrencia y auditoría.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Comprueba el permiso específico del perfil antes de realizar la acción.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_RegistrarImpresion
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@RevisionId BIGINT,@VersionOriginal BINARY(8),
    @Formato NVARCHAR(10),@VersionFormato NVARCHAR(20),@ContenidoHTML NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'PRESUPUESTOS.IMPRIMIR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF (SELECT COUNT(DISTINCT permiso.CodigoPermiso) FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso IN (N'PRESUPUESTOS.CONSULTAR',N'REPORTES.CONSULTAR'))<>2
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para generar impresiones.',1;
        IF @Formato IS NULL OR @Formato NOT IN (N'ANTERIOR',N'NUEVO') OR NULLIF(@VersionFormato,N'') IS NULL OR @ContenidoHTML IS NULL
            OR LEN(@ContenidoHTML)<100 OR LEN(@ContenidoHTML)>2000000 OR LOWER(LEFT(@ContenidoHTML,15))<>N'<!doctype html>'
            THROW 51322,N'[HYT_VALIDACION] Formato o contenido de impresión no válido.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @Estado NVARCHAR(20);
        SELECT @Estado=revision.Estado FROM dbo.PresupuestoRevision AS revision WITH (UPDLOCK,HOLDLOCK) WHERE revision.RevisionId=@RevisionId AND revision.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió mientras se generaba la impresión. Vuelve a consultarla.',1;
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@ContenidoHTML)),2);
        INSERT INTO dbo.ImpresionPresupuesto (RevisionId,Formato,VersionFormato,ContenidoHTML,SHA256,UsuarioRegistro)
        VALUES (@RevisionId,@Formato,@VersionFormato,@ContenidoHTML,@Huella,@UsuarioId);
        DECLARE @ImpresionId BIGINT=CONVERT(BIGINT,SCOPE_IDENTITY()),@Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@Cambios NVARCHAR(MAX);
        SET @Cambios=(SELECT @ImpresionId AS ImpresionId,@Formato AS Formato,@VersionFormato AS VersionFormato,@Huella AS SHA256 FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'GENERAR_IMPRESION',@Estado,@Estado,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @ImpresionId AS ImpresionId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
