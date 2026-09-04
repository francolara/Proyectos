-- =============================================
-- Author:        FRANCO LARA
-- Create date:   01/09/2026
-- Firma:         Suspende automaticamente pruebas y contratos vencidos de complejos activos, conservando su informacion comercial y excluyendo bajas.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcion_SuspenderVencidas
    @Usuario NVARCHAR(200) = NULL,
    @CantidadSuspendida INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        DECLARE @Vencidas TABLE
        (
            SuscripcionId INT NOT NULL,
            NegocioId INT NOT NULL,
            EstadoAnterior INT NOT NULL,
            EsPruebaAnterior BIT NOT NULL,
            EsPruebaConservada BIT NOT NULL,
            TipoCobro NVARCHAR(20) NULL,
            PlanComercial NVARCHAR(20) NULL,
            TipoPlan NVARCHAR(20) NULL,
            SedesPermitidas INT NULL,
            EspaciosPermitidos INT NULL,
            UsuariosPermitidos INT NULL,
            FechaInicioReferencia DATE NULL,
            FechaFinReferencia DATE NULL,
            DiasGracia INT NULL
        );

        SET @CantidadSuspendida = 0;

        BEGIN TRANSACTION;

        INSERT INTO @Vencidas
        (
            SuscripcionId, NegocioId, EstadoAnterior,
            EsPruebaAnterior, EsPruebaConservada,
            TipoCobro, PlanComercial, TipoPlan,
            SedesPermitidas, EspaciosPermitidos, UsuariosPermitidos,
            FechaInicioReferencia, FechaFinReferencia, DiasGracia
        )
        SELECT
            ns.Id,
            ns.NegocioId,
            ns.EstadoSuscripcion,
            ns.EsPrueba,
            CASE
                WHEN ns.EsPrueba = 1 THEN CAST(1 AS BIT)
                WHEN ns.EstadoSuscripcion = 3
                     AND ns.FechaFinPrueba IS NOT NULL
                     AND ns.FechaFinPlan IS NULL THEN CAST(1 AS BIT)
                ELSE CAST(0 AS BIT)
            END,
            ns.TipoCobro,
            ns.PlanComercial,
            n.TipoPlan,
            n.SedesPermitidas,
            n.EspaciosPermitidos,
            n.UsuariosPermitidos,
            CASE
                WHEN ns.EsPrueba = 1
                     OR (ns.EstadoSuscripcion = 3 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPlan IS NULL)
                    THEN ns.FechaInicioPrueba
                ELSE ns.FechaInicioPlan
            END,
            CASE
                WHEN ns.EsPrueba = 1
                     OR (ns.EstadoSuscripcion = 3 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPlan IS NULL)
                    THEN ns.FechaFinPrueba
                ELSE COALESCE(ns.FechaFinGracia, ns.FechaFinPlan)
            END,
            ns.DiasGracia
        FROM dbo.NegociosSuscripcion ns WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN dbo.Negocios n ON n.Id = ns.NegocioId
        WHERE
            n.Activo = 1
            AND
            (
              (
                ns.EstadoSuscripcion = 1
                AND ns.EsPrueba = 1
                AND ns.FechaFinPrueba IS NOT NULL
                AND ns.FechaFinPrueba < @Hoy
              )
            OR
            (
                ns.EstadoSuscripcion = 2
                AND ns.EsPrueba = 0
                AND COALESCE(ns.FechaFinGracia, ns.FechaFinPlan) IS NOT NULL
                AND COALESCE(ns.FechaFinGracia, ns.FechaFinPlan) < @Hoy
            )
            OR
            (
                ns.EstadoSuscripcion = 3
                AND
                (
                    (ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPlan IS NULL AND ns.FechaFinPrueba < @Hoy)
                    OR
                    (ns.FechaFinPlan IS NOT NULL AND COALESCE(ns.FechaFinGracia, ns.FechaFinPlan) < @Hoy)
                )
            )
            );

        UPDATE ns
        SET ns.EstadoSuscripcion = 4,
            ns.EsPrueba = v.EsPruebaConservada,
            ns.FechaActualizacion = SYSUTCDATETIME(),
            ns.UsuarioActualizacion = COALESCE(@Usuario, N'sistema')
        FROM dbo.NegociosSuscripcion ns
        INNER JOIN @Vencidas v ON v.SuscripcionId = ns.Id;

        SET @CantidadSuspendida = @@ROWCOUNT;

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
        SELECT
            v.NegocioId, v.SuscripcionId, N'SUSPENSION_AUTOMATICA',
            v.EstadoAnterior, 4,
            v.EsPruebaAnterior, v.EsPruebaConservada,
            v.TipoCobro, v.TipoCobro,
            v.PlanComercial, v.PlanComercial,
            v.TipoPlan, v.TipoPlan,
            v.SedesPermitidas, v.SedesPermitidas,
            v.EspaciosPermitidos, v.EspaciosPermitidos,
            v.UsuariosPermitidos, v.UsuariosPermitidos,
            v.FechaInicioReferencia, v.FechaFinReferencia,
            v.DiasGracia, NULL,
            N'Suspension automatica por vencimiento de la vigencia al ingresar al panel Super Admin.',
            SYSUTCDATETIME(), COALESCE(@Usuario, N'sistema')
        FROM @Vencidas v;

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
