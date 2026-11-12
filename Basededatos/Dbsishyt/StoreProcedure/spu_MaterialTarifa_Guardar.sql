-- Firma: FRANCO LARA | Create date: 10/11/2026 | Impide duplicar el tarifario directo de un material bajo el bloqueo de creación.
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   05/11/2026
-- Description:   Guarda tarifas unificadas con IGV, al menos un precio PEN/USD, estado, concurrencia y auditoría propia.
-- =============================================
-- Firma: FRANCO LARA | Create date: 07/11/2026 | Separa permisos de alta, edición y activación de tarifas.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MaterialTarifa_Guardar
    @UsuarioId NVARCHAR(128),@CorrelacionId UNIQUEIDENTIFIER,@MaterialId BIGINT,
    @CodigoTipoIGV NVARCHAR(10),@PrecioPEN DECIMAL(19,6),@PrecioUSD DECIMAL(19,6)=NULL,@Activo BIT=1,
    @Id BIGINT=NULL,@VersionOriginal BINARY(8)=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT=0;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @Id IS NULL THEN N'TARIFARIOS.CREAR' ELSE N'TARIFARIOS.MODIFICAR' END) = 0
            THROW 51700,N'[HYT_PERMISO] No tiene permiso para esta acción.',1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF EXISTS(SELECT permiso.Codigo FROM (VALUES(N'MATERIALES.CONSULTAR'),(N'TARIFARIOS.CONSULTAR'),(N'TARIFARIOS.EDITAR')) AS permiso(Codigo)
            WHERE NOT EXISTS(SELECT 1 FROM dbo.AspNetUsers AS usuario INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId=usuario.Id
                INNER JOIN dbo.RolPermiso AS acceso ON acceso.RoleId=rol.RoleId WHERE usuario.Id=@UsuarioId AND usuario.Activo=1 AND acceso.CodigoPermiso=permiso.Codigo))
            THROW 51200,N'[HYT_PERMISO] No tiene permiso para administrar el tarifario.',1;
        IF @CodigoTipoIGV IS NULL OR @CodigoTipoIGV NOT IN(N'CON_IGV',N'SIN_IGV') OR (@PrecioPEN IS NULL AND @PrecioUSD IS NULL) OR @PrecioPEN<0 OR @PrecioUSD<0 OR @Activo IS NULL
            THROW 51227,N'[HYT_VALIDACION] Revisa el tipo de IGV, estado e ingresa al menos un precio en soles o dólares.',1;
        IF @PrecioPEN<>ROUND(@PrecioPEN,2) OR @PrecioUSD<>ROUND(@PrecioUSD,2)
            THROW 51227,N'[HYT_VALIDACION] Los precios admiten hasta dos decimales.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @TarifarioId BIGINT,@VersionId BIGINT,@Anterior NVARCHAR(MAX),@Nuevo NVARCHAR(MAX),@EstadoAnterior NVARCHAR(40),@EstadoNuevo NVARCHAR(40),@CodigoSAP NVARCHAR(60);
        SELECT @CodigoSAP=material.CodigoMaterialSAP FROM dbo.ServicioMaterial AS material WITH(HOLDLOCK) WHERE material.MaterialId=@MaterialId;
        IF NOT EXISTS(SELECT 1 FROM dbo.ServicioMaterial AS material WHERE material.MaterialId=@MaterialId)
            THROW 51227,N'[HYT_VALIDACION] El material no existe.',1;
        IF @Id IS NOT NULL
        BEGIN
            SELECT @Anterior=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES),@VersionId=precio.VersionId
            FROM dbo.TarifarioPartida AS precio WITH(UPDLOCK,HOLDLOCK) INNER JOIN dbo.TarifarioVersion AS version ON version.VersionId=precio.VersionId
            INNER JOIN dbo.Tarifario AS tarifa ON tarifa.TarifarioId=version.TarifarioId
            WHERE precio.PartidaId=@Id AND precio.MaterialId=@MaterialId AND precio.VersionFila=@VersionOriginal
              AND (tarifa.EsTarifarioMaterial=1 OR version.Estado=N'BORRADOR');
            IF @Anterior IS NULL THROW 51209,N'[HYT_CONCURRENCIA] El tarifario cambió o no permite edición. Vuelve a consultarlo.',1;
            IF @Id IS NOT NULL AND @Activo <> TRY_CONVERT(BIT, JSON_VALUE(@Anterior,N'$.Activo'))
                AND dbo.fn_HYTPermiso(@UsuarioId,N'TARIFARIOS.ACTIVAR') = 0
                THROW 51700,N'[HYT_PERMISO] No tiene permiso para activar o desactivar este registro.',1;
            UPDATE dbo.TarifarioPartida SET CodigoTipoIGV=@CodigoTipoIGV,PrecioPEN=@PrecioPEN,PrecioUSD=@PrecioUSD,Activo=@Activo,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId WHERE PartidaId=@Id;
        END
        ELSE
        BEGIN
            DECLARE @Bloqueo INT;
            EXEC @Bloqueo=sys.sp_getapplock @Resource=N'HYT:TarifarioMaterialDirecto',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
            IF @Bloqueo<0 THROW 51227,N'[HYT_VALIDACION] No se pudo guardar. Intenta nuevamente.',1;
            SELECT @TarifarioId=tarifa.TarifarioId FROM dbo.Tarifario AS tarifa WITH(UPDLOCK,HOLDLOCK) WHERE tarifa.EsTarifarioMaterial=1;
            IF @TarifarioId IS NULL
            BEGIN
                DECLARE @Codigo NVARCHAR(40),@Numero BIGINT;
                WHILE 1=1
                BEGIN
                    SET @Numero=NEXT VALUE FOR dbo.SecuenciaCodigoInterno;
                    SET @Codigo=N'TAR-MAT-'+CONVERT(NVARCHAR(20),@Numero);
                    IF NOT EXISTS(SELECT 1 FROM dbo.Tarifario AS tarifa WHERE tarifa.Codigo=@Codigo) BREAK;
                END;
                INSERT INTO dbo.Tarifario(Codigo,Nombre,Activo,UsuarioRegistro,EsTarifarioMaterial) VALUES(@Codigo,N'Tarifario de materiales',1,@UsuarioId,1);
                SET @TarifarioId=CONVERT(BIGINT,SCOPE_IDENTITY());
                INSERT INTO dbo.TarifarioVersion(TarifarioId,NumeroVersion,CodigoMonedaPredeterminada,CodigoTipoIGVPredeterminado,Estado,UsuarioRegistro)
                VALUES(@TarifarioId,1,N'PEN',N'SIN_IGV',N'VIGENTE',@UsuarioId);
            END;
            SELECT @VersionId=version.VersionId FROM dbo.TarifarioVersion AS version WHERE version.TarifarioId=@TarifarioId AND version.NumeroVersion=1;
            IF @VersionId IS NULL THROW 51227,N'[HYT_VALIDACION] No se pudo preparar el tarifario.',1;
            IF EXISTS(SELECT 1 FROM dbo.TarifarioPartida AS partida WITH(UPDLOCK,HOLDLOCK) WHERE partida.VersionId=@VersionId AND partida.MaterialId=@MaterialId)
                THROW 51227,N'[HYT_VALIDACION] El material ya tiene un tarifario. Edita el registro existente.',1;
            INSERT INTO dbo.TarifarioPartida(VersionId,MaterialId,CodigoMoneda,CodigoTipoIGV,PrecioPEN,PrecioUSD,CodigoMaterialSAP,Activo,UsuarioRegistro)
            VALUES(@VersionId,@MaterialId,N'PEN',@CodigoTipoIGV,@PrecioPEN,@PrecioUSD,@CodigoSAP,@Activo,@UsuarioId);
            SET @Id=CONVERT(BIGINT,SCOPE_IDENTITY());
            INSERT INTO dbo.TarifarioPartidaZona(PartidaId,ZonaId) SELECT @Id,zona.ZonaId FROM dbo.MaterialZona AS zona WHERE zona.MaterialId=@MaterialId;
        END;
        SELECT @Nuevo=(SELECT precio.CodigoTipoIGV,precio.PrecioPEN,precio.PrecioUSD,precio.Activo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES)
        FROM dbo.TarifarioPartida AS precio WHERE precio.PartidaId=@Id;
        SET @EstadoAnterior=JSON_VALUE(@Anterior,N'$.Activo');SET @EstadoNuevo=JSON_VALUE(@Nuevo,N'$.Activo');
        DECLARE @Cambios NVARCHAR(MAX)=(SELECT JSON_QUERY(@Anterior) AS Anterior,JSON_QUERY(@Nuevo) AS Nuevo FOR JSON PATH,WITHOUT_ARRAY_WRAPPER);
        DECLARE @Accion NVARCHAR(60)=CASE WHEN @Anterior IS NULL THEN N'CREAR' WHEN @EstadoAnterior<>@EstadoNuevo THEN N'CAMBIAR_ESTADO' ELSE N'ACTUALIZAR' END;
        DECLARE @Registro NVARCHAR(100)=CONVERT(NVARCHAR(100),@Id);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TarifarioPartida',@Registro,@Accion,@EstadoAnterior,@EstadoNuevo,NULL,NULL,@Cambios;
        COMMIT TRANSACTION;
        SELECT @Id AS Id;
    END TRY
    BEGIN CATCH
        IF @Propia=1 AND XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000),@ErrorSeverity INT,@ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(),@ErrorSeverity=ERROR_SEVERITY(),@ErrorState=ERROR_STATE();
        RAISERROR(@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH;
END;
GO
