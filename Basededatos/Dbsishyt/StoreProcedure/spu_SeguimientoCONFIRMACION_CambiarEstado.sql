-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Cambia estado con dependencias, fecha, responsable, auditoría e idempotencia.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Comprueba el permiso específico del perfil antes de realizar la acción.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_SeguimientoCONFIRMACION_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@Id BIGINT,@PadreId BIGINT,@VersionPadre BINARY(8),@VersionOriginal BINARY(8),
    @NuevoEstado NVARCHAR(40),@FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'CONFIRMACIONES.MODIFICAR') = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'CONFIRMACIONES.EDITAR')
            THROW 51400,N'[HYT_PERMISO] No tiene permiso para cambiar este estado.',1;
        IF @@TRANCOUNT<>0 THROW 51402,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        SET @Responsable=NULLIF(LTRIM(RTRIM(@Responsable)),N'');SET @Motivo=NULLIF(LTRIM(RTRIM(@Motivo)),N'');
        IF @Id IS NULL OR @Id<=0 OR @PadreId IS NULL OR @PadreId<=0 OR @VersionPadre IS NULL OR @VersionOriginal IS NULL OR @FechaEvento IS NULL OR @Responsable IS NULL OR @CorrelacionId IS NULL OR @CorrelacionId='00000000-0000-0000-0000-000000000000'
            THROW 51430,N'[HYT_VALIDACION] Completa la fecha, responsable y referencia del registro.',1;
        DECLARE @Solicitud NVARCHAR(MAX)=(SELECT @Id AS Id,@PadreId AS PadreId,@VersionPadre AS VersionPadre,@VersionOriginal AS VersionOriginal,@NuevoEstado AS NuevoEstado,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Motivo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        DECLARE @Huella CHAR(64)=CONVERT(CHAR(64),HASHBYTES('SHA2_256',CONVERT(VARBINARY(MAX),@Solicitud)),2),@Tipo NVARCHAR(40)=N'CONFIRMACION_ESTADO';
        BEGIN TRANSACTION;SET @Propia=1;
        DECLARE @Resultado BIGINT,@HuellaAnterior CHAR(64),@ActorAnterior NVARCHAR(128),@TipoAnterior NVARCHAR(40),@PadreAnterior BIGINT;
        SELECT @Resultado=operacion.RegistroId,@HuellaAnterior=operacion.Huella,@ActorAnterior=operacion.UsuarioId,@TipoAnterior=operacion.TipoOperacion,@PadreAnterior=operacion.PadreId
        FROM dbo.DocumentoOperacion AS operacion WITH(UPDLOCK,HOLDLOCK) WHERE operacion.OperacionId=@CorrelacionId;
        IF @Resultado IS NOT NULL
        BEGIN
            IF @Resultado<>@Id OR @HuellaAnterior<>@Huella OR @ActorAnterior<>@UsuarioId OR @TipoAnterior<>@Tipo OR @PadreAnterior<>@PadreId
                THROW 51409,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos. Vuelve a abrir el registro.',1;
            COMMIT;SET @Propia=0;SELECT @Resultado AS Id;RETURN;
        END;
        DECLARE @PresupuestoId BIGINT=(SELECT oc.PresupuestoId FROM dbo.OrdenCompraCliente AS oc WHERE oc.OrdenCompraId=@PadreId),@TipoAtencion NVARCHAR(20),@EstadoPresupuesto NVARCHAR(20);
        SELECT @TipoAtencion=p.TipoAtencion FROM dbo.Presupuesto AS p WITH(UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId;
        IF @TipoAtencion IS NULL THROW 51410,N'[HYT_VALIDACION] El presupuesto no existe.',1;
        IF NOT EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH(UPDLOCK,HOLDLOCK) WHERE oc.OrdenCompraId=@PadreId AND oc.VersionFila=@VersionPadre)
            THROW 51411,N'[HYT_CONCURRENCIA] La OC cambió. Vuelve a consultarla.',1;
        SELECT TOP(1) @EstadoPresupuesto=r.Estado FROM dbo.PresupuestoRevision AS r WITH(HOLDLOCK) WHERE r.PresupuestoId=@PresupuestoId ORDER BY r.NumeroRevision DESC;
        IF @EstadoPresupuesto=N'CANCELADO' AND @NuevoEstado NOT IN(N'CANCELADA',N'ANULADA',N'RETIRADA') THROW 51412,N'[HYT_VALIDACION] El presupuesto está cancelado.',1;
        IF EXISTS(SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.ReferenciaComprobanteExterno AS referencia ON referencia.OrdenCompraId=oc.OrdenCompraId AND referencia.Vigente=1 WHERE oc.PresupuestoId=@PresupuestoId)
            THROW 51412,N'[HYT_VALIDACION] Retira primero el marcado facturado antes de cambiar documentos.',1;
        DECLARE @EstadoAnterior NVARCHAR(40),@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX);
        SELECT @EstadoAnterior=CASE WHEN documento.Vigente=1 THEN N'VIGENTE' ELSE N'RETIRADA' END,@Anterior=(SELECT documento.FechaConfirmacion,documento.NumeroConfirmacion,documento.RepresentanteCliente,documento.Observacion,documento.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.ConfirmacionFacturacion AS documento WITH(UPDLOCK,HOLDLOCK) WHERE documento.ConfirmacionId=@Id AND documento.OrdenCompraId=@PadreId AND documento.VersionFila=@VersionOriginal;
        IF @Anterior IS NULL THROW 51411,N'[HYT_CONCURRENCIA] El documento cambió o no pertenece al expediente. Vuelve a consultarlo.',1;
        IF NOT((@EstadoAnterior=N'VIGENTE' AND @NuevoEstado=N'RETIRADA')) OR @NuevoEstado IS NULL THROW 51431,N'[HYT_VALIDACION] Transición de estado no permitida.',1;
        IF @Motivo IS NULL AND (@NuevoEstado IN(N'OBSERVADA',N'ANULADA',N'CANCELADA',N'RETIRADA',N'PENDIENTE') OR @EstadoAnterior IN(N'ANULADA',N'CANCELADA',N'ACEPTADA',N'TERMINADA',N'CERRADA_OPERATIVAMENTE') OR (@EstadoAnterior=N'TRABAJO_COMPLETADO' AND @NuevoEstado=N'REGISTRADA'))
            THROW 51432,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        UPDATE dbo.ConfirmacionFacturacion SET Vigente=0,FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE ConfirmacionId=@Id;
        UPDATE dbo.OrdenCompraCliente SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE OrdenCompraId=@PadreId;
        SELECT @Nuevo=(SELECT actual.FechaConfirmacion,actual.NumeroConfirmacion,actual.RepresentanteCliente,actual.Observacion,actual.Vigente FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES) FROM dbo.ConfirmacionFacturacion AS actual WHERE actual.ConfirmacionId=@Id;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id),@Detalle NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo,@Responsable AS Responsable FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'ConfirmacionFacturacion',@Registro,N'CAMBIAR_ESTADO',@EstadoAnterior,@NuevoEstado,@FechaEvento,@Motivo,@Detalle;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        INSERT INTO dbo.DocumentoOperacion(OperacionId,TipoOperacion,PadreId,RegistroId,Huella,UsuarioId) VALUES(@CorrelacionId,@Tipo,@PadreId,@Id,@Huella,@UsuarioId);
        COMMIT;SET @Propia=0;SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
