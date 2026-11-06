-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Transiciones auditadas y regularización de emergencia con OC sin modificar contenido.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_Presupuesto_CambiarEstado
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@PresupuestoId BIGINT,@RevisionId BIGINT,
    @VersionOriginal BINARY(8),@VersionPresupuesto BINARY(8),@NuevoEstado NVARCHAR(20),
    @FechaEvento DATE,@Responsable NVARCHAR(150),@Motivo NVARCHAR(2000)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF @@TRANCOUNT<>0 THROW 51302,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF NOT EXISTS (SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId=rol.RoleId
            WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND permiso.CodigoPermiso=N'PRESUPUESTOS.EDITAR')
            THROW 51300,N'[HYT_PERMISO] No tiene permiso para cambiar el estado.',1;
        IF @FechaEvento IS NULL OR NULLIF(LTRIM(RTRIM(@Responsable)),N'') IS NULL OR @NuevoEstado IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica fecha y responsable de la acción.',1;
        IF @NuevoEstado IN (N'OBSERVADO',N'RECHAZADO',N'CANCELADO',N'VENCIDO') AND NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL
            THROW 51318,N'[HYT_VALIDACION] Indica el motivo del cambio de estado.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        IF NOT EXISTS (SELECT 1 FROM dbo.Presupuesto AS p WITH (UPDLOCK,HOLDLOCK) WHERE p.PresupuestoId=@PresupuestoId AND p.VersionFila=@VersionPresupuesto)
            THROW 51309,N'[HYT_CONCURRENCIA] El presupuesto cambió. Vuelve a consultarlo.',1;
        DECLARE @Estado NVARCHAR(20),@Empresa NVARCHAR(MAX),@Moneda NVARCHAR(3),@Total DECIMAL(19,2);
        SELECT @Estado=r.Estado,@Empresa=r.EmpresaSnapshotJson,@Moneda=r.CodigoMoneda,@Total=r.Total
        FROM dbo.PresupuestoRevision AS r WITH (UPDLOCK,HOLDLOCK)
        WHERE r.RevisionId=@RevisionId AND r.PresupuestoId=@PresupuestoId AND r.VersionFila=@VersionOriginal;
        IF @Estado IS NULL THROW 51309,N'[HYT_CONCURRENCIA] La revisión cambió o no pertenece al presupuesto.',1;
        IF EXISTS (SELECT 1 FROM dbo.PresupuestoRevision AS r WHERE r.PresupuestoId=@PresupuestoId AND r.NumeroRevision>(SELECT origen.NumeroRevision FROM dbo.PresupuestoRevision AS origen WHERE origen.RevisionId=@RevisionId))
            THROW 51319,N'[HYT_VALIDACION] Solo se cambia el estado de la última revisión.',1;
        IF NOT ((@Estado=N'BORRADOR' AND @NuevoEstado IN (N'ENVIADO',N'CANCELADO'))
            OR (@Estado=N'ENVIADO' AND @NuevoEstado IN (N'OBSERVADO',N'APROBADO',N'RECHAZADO',N'VENCIDO',N'CANCELADO'))
            OR (@Estado IN (N'OBSERVADO',N'RECHAZADO',N'VENCIDO',N'APROBADO') AND @NuevoEstado=N'CANCELADO'))
            THROW 51319,N'[HYT_VALIDACION] La transición no está permitida. Para corregir contenido crea otra revisión.',1;
        IF EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc WITH (HOLDLOCK) WHERE oc.PresupuestoId=@PresupuestoId)
            AND NOT EXISTS (SELECT 1 FROM dbo.OrdenCompraCliente AS oc INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=oc.PresupuestoId
                WHERE oc.PresupuestoId=@PresupuestoId AND oc.RevisionId=@RevisionId AND oc.Estado<>N'CANCELADA' AND p.TipoAtencion=N'EMERGENCIA'
                AND ((@Estado=N'BORRADOR' AND @NuevoEstado=N'ENVIADO') OR (@Estado=N'ENVIADO' AND @NuevoEstado=N'APROBADO')))
            THROW 51320,N'[HYT_VALIDACION] La revisión vinculada a una OC no permite este cambio.',1;
        IF @NuevoEstado=N'CANCELADO' AND EXISTS (SELECT 1 FROM dbo.GuiaConformidad AS guia WITH (HOLDLOCK) WHERE guia.PresupuestoId=@PresupuestoId AND guia.Estado<>N'ANULADA')
            THROW 51320,N'[HYT_VALIDACION] No se puede cancelar mientras exista una guía vigente.',1;
        IF @NuevoEstado=N'ENVIADO'
        BEGIN
            IF @Total IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId)
                OR EXISTS (SELECT 1 FROM dbo.PresupuestoPartida AS linea WHERE linea.RevisionId=@RevisionId AND (linea.Total IS NULL OR linea.PrecioUnitario IS NULL OR linea.CodigoMonedaOrigen<>@Moneda))
                THROW 51321,N'[HYT_VALIDACION] Completa las partidas y los precios en la moneda elegida antes de enviar.',1;
            IF ISJSON(@Empresa)<>1 OR @Empresa IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.Ruc'),N'') IS NULL OR NULLIF(JSON_VALUE(@Empresa,N'$.RazonSocial'),N'') IS NULL
                THROW 51321,N'[HYT_VALIDACION] Configura los datos de HYT y vuelve a guardar el borrador antes de enviar.',1;
        END;
        UPDATE dbo.PresupuestoRevision SET Estado=@NuevoEstado,FechaAprobacion=CASE WHEN @NuevoEstado=N'APROBADO' THEN @FechaEvento ELSE FechaAprobacion END,
            FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE RevisionId=@RevisionId;
        UPDATE dbo.Presupuesto SET FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PresupuestoId=@PresupuestoId;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@RevisionId),@FechaReal DATETIME2(3)=CONVERT(DATETIME2(3),@FechaEvento),
            @MotivoAudit NVARCHAR(1000)=LEFT(@Motivo,1000),@Detalle NVARCHAR(MAX);
        SET @Detalle=(SELECT @Estado AS EstadoAnterior,@NuevoEstado AS EstadoNuevo,@FechaEvento AS FechaEvento,@Responsable AS Responsable,@Motivo AS Observacion FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'PresupuestoRevision',@Registro,N'CAMBIAR_ESTADO',@Estado,@NuevoEstado,@FechaReal,@MotivoAudit,@Detalle;
        COMMIT TRANSACTION;
        SELECT p.PresupuestoId,r.RevisionId,p.VersionFila AS VersionPresupuesto,r.VersionFila AS VersionRevision,r.Estado
        FROM dbo.Presupuesto AS p INNER JOIN dbo.PresupuestoRevision AS r ON r.PresupuestoId=p.PresupuestoId AND r.RevisionId=@RevisionId WHERE p.PresupuestoId=@PresupuestoId;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
