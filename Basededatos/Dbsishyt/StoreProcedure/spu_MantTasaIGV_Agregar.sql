-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Agrega vigencia IGV con concurrencia y auditoría del alta y cierre anterior.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_MantTasaIGV_Agregar
    @UsuarioId NVARCHAR(128), @CorrelacionId UNIQUEIDENTIFIER, @FechaDesde DATE, @Valor DECIMAL(9,6),
    @Motivo NVARCHAR(1000), @UltimaFechaEsperada DATE=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    DECLARE @Propia BIT = 0;
    BEGIN TRY
        IF NOT EXISTS
        (
            SELECT 1 FROM dbo.AspNetUsers AS usuario
            INNER JOIN dbo.AspNetUserRoles AS rol ON rol.UserId = usuario.Id
            INNER JOIN dbo.RolPermiso AS permiso ON permiso.RoleId = rol.RoleId
            WHERE usuario.Id = @UsuarioId AND usuario.Activo = 1 AND permiso.CodigoPermiso = N'CONFIGURACION.EDITAR'
        ) THROW 51200, N'[HYT_PERMISO] No tiene permiso para esta operación.', 1;
        IF @@TRANCOUNT<>0 THROW 51202,N'[HYT_VALIDACION] La operación administra su transacción.',1;
        IF @FechaDesde IS NULL OR @Valor IS NULL OR @Valor<0 OR @Valor>1 OR NULLIF(LTRIM(RTRIM(@Motivo)),N'') IS NULL
            THROW 51226,N'[HYT_VALIDACION] Indica fecha, tasa entre cero y uno y motivo.',1;
        BEGIN TRANSACTION;
        SET @Propia=1;
        DECLARE @UltimaFecha DATE,@AnteriorValor DECIMAL(9,6),@AnteriorHasta DATE;
        SELECT @UltimaFecha=MAX(tasa.FechaDesde) FROM dbo.TasaIGV AS tasa WITH (UPDLOCK,HOLDLOCK);
        IF (@UltimaFecha<>@UltimaFechaEsperada) OR (@UltimaFecha IS NULL AND @UltimaFechaEsperada IS NOT NULL) OR (@UltimaFecha IS NOT NULL AND @UltimaFechaEsperada IS NULL)
            THROW 51209,N'[HYT_CONCURRENCIA] Las vigencias cambiaron. Actualiza el listado antes de registrar otra.',1;
        IF @UltimaFecha IS NOT NULL AND @FechaDesde<=@UltimaFecha THROW 51227,N'[HYT_VALIDACION] La nueva vigencia debe ser posterior a la última registrada.',1;
        SELECT @AnteriorValor=tasa.Valor,@AnteriorHasta=tasa.FechaHasta FROM dbo.TasaIGV AS tasa WHERE tasa.FechaDesde=@UltimaFecha;
        UPDATE dbo.TasaIGV SET FechaHasta=@FechaDesde WHERE FechaDesde=@UltimaFecha;
        INSERT INTO dbo.TasaIGV (FechaDesde,FechaHasta,Valor) VALUES (@FechaDesde,NULL,@Valor);
        DECLARE @Registro NVARCHAR(100),@Detalle NVARCHAR(MAX);
        IF @UltimaFecha IS NOT NULL
        BEGIN
            SET @Registro=CONVERT(NVARCHAR(10),@UltimaFecha,23);
            SET @Detalle=(SELECT @AnteriorValor AS Valor,@AnteriorHasta AS FechaHastaAnterior,@FechaDesde AS FechaHastaNueva FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TasaIGV',@Registro,N'CERRAR_VIGENCIA',NULL,NULL,NULL,@Motivo,@Detalle;
        END;
        SET @Registro=CONVERT(NVARCHAR(10),@FechaDesde,23);
        SET @Detalle=(SELECT @Valor AS Valor,@FechaDesde AS FechaDesde,@UltimaFecha AS VigenciaAnterior FOR JSON PATH,WITHOUT_ARRAY_WRAPPER,INCLUDE_NULL_VALUES);
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@CorrelacionId,N'TasaIGV',@Registro,N'NUEVA_VIGENCIA',NULL,NULL,NULL,@Motivo,@Detalle;
        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF @Propia = 1 AND XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
