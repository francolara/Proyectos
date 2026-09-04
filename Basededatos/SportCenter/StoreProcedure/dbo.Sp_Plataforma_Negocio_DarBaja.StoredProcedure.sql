-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/09/2026
-- Firma:         Da de baja logicamente un complejo, suspende su servicio y conserva todos sus datos e historial.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocio_DarBaja
    @NegocioId INT,
    @Motivo NVARCHAR(100),
    @Observacion NVARCHAR(500) = NULL,
    @ConfirmacionNombre NVARCHAR(200),
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @NombreComercial NVARCHAR(200);
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
        SET @ConfirmacionNombre = NULLIF(LTRIM(RTRIM(@ConfirmacionNombre)), N'');

        IF @Motivo IS NULL
            RAISERROR('Debes indicar el motivo de la baja.', 16, 1);

        SELECT
            @NombreComercial = n.NombreComercial,
            @TipoPlan = n.TipoPlan,
            @SedesPermitidas = n.SedesPermitidas,
            @EspaciosPermitidos = n.EspaciosPermitidos,
            @UsuariosPermitidos = n.UsuariosPermitidos
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @NombreComercial IS NULL
            RAISERROR('El complejo deportivo no existe o ya fue dado de baja.', 16, 1);

        IF @ConfirmacionNombre IS NULL OR @ConfirmacionNombre COLLATE Latin1_General_100_CI_AI <> @NombreComercial COLLATE Latin1_General_100_CI_AI
            RAISERROR('El nombre de confirmacion no coincide con el complejo deportivo.', 16, 1);

        SELECT
            @SuscripcionId = ns.Id,
            @EstadoAnterior = ns.EstadoSuscripcion,
            @EsPrueba = ns.EsPrueba,
            @TipoCobro = ns.TipoCobro,
            @PlanComercial = ns.PlanComercial,
            @FechaInicioReferencia = CASE WHEN ns.EsPrueba = 1 THEN ns.FechaInicioPrueba ELSE ns.FechaInicioPlan END,
            @FechaFinReferencia = CASE WHEN ns.EsPrueba = 1 THEN ns.FechaFinPrueba ELSE COALESCE(ns.FechaFinGracia, ns.FechaFinPlan) END,
            @DiasGracia = ns.DiasGracia
        FROM dbo.NegociosSuscripcion ns
        WHERE ns.NegocioId = @NegocioId;

        IF @SuscripcionId IS NULL
            RAISERROR('No se encontro la suscripcion del complejo deportivo.', 16, 1);

        SET @DetalleMovimiento = LEFT(
            N'Motivo: ' + @Motivo
            + CASE WHEN @Observacion IS NULL THEN N'' ELSE N'. ' + @Observacion END,
            500);

        BEGIN TRANSACTION;

        UPDATE dbo.Negocios
        SET Activo = 0
        WHERE Id = @NegocioId
          AND Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se pudo dar de baja el complejo deportivo.', 16, 1);

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
            @NegocioId, @SuscripcionId, N'BAJA_COMPLEJO',
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
