-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/09/2026
-- Firma:         Reactiva un servicio suspendido cuando la prueba o el contrato conservan vigencia comercial.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_Reactivar
    @NegocioId INT,
    @Observacion NVARCHAR(500) = NULL,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);
        DECLARE @SuscripcionId INT;
        DECLARE @EstadoAnterior INT;
        DECLARE @EstadoNuevo INT;
        DECLARE @EsPrueba BIT;
        DECLARE @FechaInicioPrueba DATE;
        DECLARE @FechaFinPrueba DATE;
        DECLARE @TipoCobro NVARCHAR(20);
        DECLARE @FechaInicioPlan DATE;
        DECLARE @FechaFinPlan DATE;
        DECLARE @PlanComercial NVARCHAR(20);
        DECLARE @TipoPlan NVARCHAR(20);
        DECLARE @SedesPermitidas INT;
        DECLARE @EspaciosPermitidos INT;
        DECLARE @UsuariosPermitidos INT;
        DECLARE @DiasGracia INT;

        SET @Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), N'');

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPrueba = ns.EsPrueba,
            @FechaInicioPrueba = ns.FechaInicioPrueba,
            @FechaFinPrueba = ns.FechaFinPrueba,
            @TipoCobro = ns.TipoCobro,
            @FechaInicioPlan = ns.FechaInicioPlan,
            @FechaFinPlan = ns.FechaFinPlan,
            @PlanComercial = ns.PlanComercial,
            @TipoPlan = n.TipoPlan,
            @SedesPermitidas = n.SedesPermitidas,
            @EspaciosPermitidos = n.EspaciosPermitidos,
            @UsuariosPermitidos = n.UsuariosPermitidos,
            @DiasGracia = ns.DiasGracia
        FROM dbo.NegociosSuscripcion ns
        INNER JOIN dbo.Negocios n ON n.Id = ns.NegocioId
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('No se encontro una suscripcion para el complejo deportivo.', 16, 1);

        IF @EstadoAnterior <> 4
            RAISERROR('El servicio del complejo deportivo no esta suspendido.', 16, 1);

        IF @EsPrueba = 1
        BEGIN
            IF @FechaFinPrueba IS NULL OR @FechaFinPrueba < @Hoy
                RAISERROR('La prueba ya vencio. Debes extenderla antes de reactivar el servicio.', 16, 1);

            SET @EstadoNuevo = 1;
        END
        ELSE
        BEGIN
            IF @TipoCobro IS NULL OR @FechaFinPlan IS NULL OR @FechaFinPlan < @Hoy
                RAISERROR('El contrato ya vencio. Debes asignar una nueva vigencia antes de reactivar el servicio.', 16, 1);

            SET @EstadoNuevo = 2;
        END;

        BEGIN TRANSACTION;

        UPDATE dbo.NegociosSuscripcion
        SET EstadoSuscripcion = @EstadoNuevo,
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
            @NegocioId, @SuscripcionId, N'REACTIVACION',
            @EstadoAnterior, @EstadoNuevo,
            @EsPrueba, @EsPrueba,
            @TipoCobro, @TipoCobro,
            @PlanComercial, @PlanComercial,
            @TipoPlan, @TipoPlan,
            @SedesPermitidas, @SedesPermitidas,
            @EspaciosPermitidos, @EspaciosPermitidos,
            @UsuariosPermitidos, @UsuariosPermitidos,
            CASE WHEN @EsPrueba = 1 THEN @FechaInicioPrueba ELSE @FechaInicioPlan END,
            CASE WHEN @EsPrueba = 1 THEN @FechaFinPrueba ELSE @FechaFinPlan END,
            @DiasGracia, NULL, COALESCE(@Observacion, N'Reactivacion manual del servicio desde superadmin.'),
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
