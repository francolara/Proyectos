-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/09/2026
-- Firma:         Reactiva logicamente un complejo dado de baja y mantiene su servicio suspendido hasta validar la vigencia.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocio_Reactivar
    @NegocioId INT,
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

        SET @Observacion = NULLIF(LTRIM(RTRIM(@Observacion)), N'');

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
            @FechaFinReferencia = CASE WHEN ns.EsPrueba = 1 THEN ns.FechaFinPrueba ELSE COALESCE(ns.FechaFinGracia, ns.FechaFinPlan) END,
            @DiasGracia = ns.DiasGracia
        FROM dbo.Negocios n
        INNER JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE n.Id = @NegocioId
          AND n.Activo = 0;

        IF @SuscripcionId IS NULL
            RAISERROR('El complejo deportivo no existe o no se encuentra dado de baja.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Negocios
        SET Activo = 1
        WHERE Id = @NegocioId
          AND Activo = 0;

        IF @@ROWCOUNT = 0
            RAISERROR('No se pudo reactivar el complejo deportivo.', 16, 1);

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
            @NegocioId, @SuscripcionId, N'REACTIVACION_COMPLEJO',
            @EstadoAnterior, 4,
            @EsPrueba, @EsPrueba,
            @TipoCobro, @TipoCobro,
            @PlanComercial, @PlanComercial,
            @TipoPlan, @TipoPlan,
            @SedesPermitidas, @SedesPermitidas,
            @EspaciosPermitidos, @EspaciosPermitidos,
            @UsuariosPermitidos, @UsuariosPermitidos,
            @FechaInicioReferencia, @FechaFinReferencia,
            @DiasGracia, NULL, COALESCE(@Observacion, N'Reactivacion logica del complejo desde superadmin. El servicio permanece suspendido.'),
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
