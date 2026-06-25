-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra o actualiza la configuracion contable automatica por empresa, modulo y escenario.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   22/06/2026
-- Description:   Aclara el mensaje de validacion para exigir cuentas de la empresa, activas y con movimiento.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarConfiguracionContabilizacion
    @IdConfiguracionContabilizacion INT = NULL,
    @IdEmpresa INT,
    @ModuloOperacion VARCHAR(10),
    @EscenarioOperacion VARCHAR(20),
    @IdOrigen INT,
    @Descripcion NVARCHAR(200),
    @GeneraAsientoAutomatico BIT,
    @UsaTipoCambio BIT,
    @Activo BIT,
    @DetalleXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdConfiguracionTrabajo INT

        IF @DetalleXml IS NULL
        BEGIN
            RAISERROR(N'Debe enviar el detalle de la configuracion contable.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdOrigen = @IdOrigen
              AND o.IdEmpresa = @IdEmpresa
              AND o.Estado = 1
        )
        BEGIN
            RAISERROR(N'El origen indicado no existe o no pertenece a la empresa.', 16, 1);
        END;

        DECLARE @Detalle TABLE
        (
            Orden SMALLINT NOT NULL,
            ComponenteContable VARCHAR(20) NOT NULL,
            IdPlanCuenta INT NOT NULL,
            NaturalezaMovimiento CHAR(1) NOT NULL,
            Activo BIT NOT NULL
        );

        INSERT INTO @Detalle
        (
            Orden,
            ComponenteContable,
            IdPlanCuenta,
            NaturalezaMovimiento,
            Activo
        )
        SELECT
            T.N.value('@Orden', 'smallint'),
            T.N.value('@ComponenteContable', 'varchar(20)'),
            T.N.value('@IdPlanCuenta', 'int'),
            T.N.value('@NaturalezaMovimiento', 'char(1)'),
            T.N.value('@Activo', 'bit')
        FROM @DetalleXml.nodes('/Detalles/Detalle') AS T(N);

        IF NOT EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
        )
        BEGIN
            RAISERROR(N'Debe registrar al menos un componente contable.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT d.ComponenteContable
            FROM @Detalle AS d
            WHERE d.Activo = 1
            GROUP BY
                d.ComponenteContable
            HAVING COUNT(1) > 1
        )
        BEGIN
            RAISERROR(N'No se permiten componentes activos duplicados en la misma configuracion.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Detalle AS d
            LEFT JOIN dbo.CON_PlanCuenta AS p
                ON p.IdPlanCuenta = d.IdPlanCuenta
               AND p.IdEmpresa = @IdEmpresa
               AND p.Estado = 1
               AND p.AceptaMovimiento = 1
            WHERE p.IdPlanCuenta IS NULL
        )
        BEGIN
            RAISERROR(N'Todas las cuentas configuradas deben existir, estar activas y aceptar movimiento.', 16, 1);
        END;

        BEGIN TRAN;

        IF @IdConfiguracionContabilizacion IS NULL
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_ConfiguracionContabilizacion AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.ModuloOperacion = @ModuloOperacion
                  AND c.EscenarioOperacion = @EscenarioOperacion
            )
            BEGIN
                RAISERROR(N'Ya existe una configuracion para la empresa, modulo y escenario seleccionados.', 16, 1);
            END;

            INSERT INTO dbo.CON_ConfiguracionContabilizacion
            (
                IdEmpresa,
                ModuloOperacion,
                EscenarioOperacion,
                IdOrigen,
                Descripcion,
                GeneraAsientoAutomatico,
                UsaTipoCambio,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @ModuloOperacion,
                @EscenarioOperacion,
                @IdOrigen,
                @Descripcion,
                @GeneraAsientoAutomatico,
                @UsaTipoCambio,
                @Activo,
                @UsuarioRegistro
            );

            SET @IdConfiguracionTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            SET @IdConfiguracionTrabajo = @IdConfiguracionContabilizacion;

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_ConfiguracionContabilizacion AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.ModuloOperacion = @ModuloOperacion
                  AND c.EscenarioOperacion = @EscenarioOperacion
                  AND c.IdConfiguracionContabilizacion <> @IdConfiguracionContabilizacion
            )
            BEGIN
                RAISERROR(N'Ya existe otra configuracion para la empresa, modulo y escenario seleccionados.', 16, 1);
            END;

            UPDATE dbo.CON_ConfiguracionContabilizacion
            SET ModuloOperacion = @ModuloOperacion,
                EscenarioOperacion = @EscenarioOperacion,
                IdOrigen = @IdOrigen,
                Descripcion = @Descripcion,
                GeneraAsientoAutomatico = @GeneraAsientoAutomatico,
                UsaTipoCambio = @UsaTipoCambio,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion
              AND IdEmpresa = @IdEmpresa;

            IF @@ROWCOUNT = 0
            BEGIN
                RAISERROR(N'La configuracion indicada no existe para la empresa activa.', 16, 1);
            END;

            DELETE FROM dbo.CON_ConfiguracionContabilizacionDetalle
            WHERE IdConfiguracionContabilizacion = @IdConfiguracionContabilizacion;
        END;

        INSERT INTO dbo.CON_ConfiguracionContabilizacionDetalle
        (
            IdConfiguracionContabilizacion,
            Orden,
            ComponenteContable,
            IdPlanCuenta,
            NaturalezaMovimiento,
            Activo,
            UsuarioRegistro
        )
        SELECT
            @IdConfiguracionTrabajo,
            d.Orden,
            d.ComponenteContable,
            d.IdPlanCuenta,
            d.NaturalezaMovimiento,
            d.Activo,
            @UsuarioRegistro
        FROM @Detalle AS d
        ORDER BY
            d.Orden ASC;

        COMMIT;

        SELECT
            @IdConfiguracionTrabajo AS IdConfiguracionContabilizacion;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)

    END CATCH

END
