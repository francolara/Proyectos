-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/09/2026
-- Firma:         Suspende temporalmente el servicio conservando prueba, plan, vigencia y limites para una posible reactivacion.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_Suspender
    @NegocioId INT,
    @Motivo NVARCHAR(100),
    @Observacion NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EsPrueba BIT;
        DECLARE @TipoCobro NVARCHAR(20);
        DECLARE @PlanComercial NVARCHAR(20);
        DECLARE @TipoPlan NVARCHAR(20);
        DECLARE @SedesPermitidas INT;
        DECLARE @EspaciosPermitidos INT;
        DECLARE @UsuariosPermitidos INT;
        DECLARE @FechaInicioReferencia DATE;
        DECLARE @FechaFinReferencia DATE;
        DECLARE @DiasGracia INT;
        DECLARE @DetalleMovimiento NVARCHAR(500);

        SET @Motivo = NULLIF(LTRIM(RTRIM(@Motivo)), N'');
        SET @Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), N'');

        IF @Motivo IS NULL
            RAISERROR('Debes indicar el motivo de la suspension.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPrueba = ns.EsPrueba,
            @TipoCobro = ns.TipoCobro,
            @PlanComercial = ns.PlanComercial,
            @TipoPlan = n.TipoPlan,
            @SedesPermitidas = n.SedesPermitidas,
            @EspaciosPermitidos = n.EspaciosPermitidos,
            @UsuariosPermitidos = n.UsuariosPermitidos,
            @FechaInicioReferencia = CASE WHEN ns.EsPrueba = 1 THEN ns.FechaInicioPrueba ELSE ns.FechaInicioPlan END,
            @FechaFinReferencia = CASE WHEN ns.EsPrueba = 1 THEN ns.FechaFinPrueba ELSE ns.FechaFinPlan END,
            @DiasGracia = ns.DiasGracia
        FROM dbo.NegociosSuscripcion ns
        INNER JOIN dbo.Negocios n ON n.Id = ns.NegocioId
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('No se encontro una suscripcion para el complejo deportivo.', 16, 1);

        IF @EstadoAnterior = 4
            RAISERROR('El servicio del complejo deportivo ya esta suspendido.', 16, 1);

        IF @EstadoAnterior NOT IN (1, 2)
            RAISERROR('Solo se puede suspender una prueba o contrato activo.', 16, 1);

        SET @DetalleMovimiento = LEFT(
            N'Motivo: ' + @Motivo
            + CASE WHEN @Observacion IS NULL THEN N'' ELSE N'. ' + @Observacion END,
            500);

        BEGIN TRANSACTION;

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = 4,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        WHERE Id = @SuscripcionId;

        INSERT INTO dbo.NegociosSuscripcionMovimiento
        (
            NegocioId, NegocioSuscripcionId, TipoMovimiento,
            EstadoSuscripcionAnterior, EstadoSuscripcionNuevo,
            EsPruebaAnterior, EsPruebaNuevo,
            TipoCobroAnterior, TipoCobroNuevo,
            PlanComercialAnterior, PlanComercialNuevo,
            TipoPlanAnterior, TipoPlanNuevo,
            SedesPermitidasAnterior, SedesPermitidasNuevo,
            EspaciosPermitidosAnterior, EspaciosPermitidosNuevo,
            UsuariosPermitidosAnterior, UsuariosPermitidosNuevo,
            FechaInicioReferencia, FechaFinReferencia,
            DiasGracia, DiasExtra, Observacion,
            FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @SuscripcionId, N'SUSPENSION',
            @EstadoAnterior, 4,
            @EsPrueba, @EsPrueba,
            @TipoCobro, @TipoCobro,
            @PlanComercial, @PlanComercial,
            @TipoPlan, @TipoPlan,
            @SedesPermitidas, @SedesPermitidas,
            @EspaciosPermitidos, @EspaciosPermitidos,
            @UsuariosPermitidos, @UsuariosPermitidos,
            @FechaInicioReferencia, @FechaFinReferencia,
            @DiasGracia, NULL, @DetalleMovimiento,
            SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
        );

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
    END CATCH
END
