-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Crea o actualiza una cuenta maestra conservando inmutable su codigo cuando ya existe.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarPlanCuentaMaestro
    @IdPlanCuentaMaestro INT = NULL,
    @CodigoCuenta VARCHAR(20),
    @CodigoCuentaPadre VARCHAR(20) = NULL,
    @NombreCuenta NVARCHAR(200),
    @ColBalance CHAR(1),
    @IdMoneda VARCHAR(3) = '',
    @TipoCambio CHAR(1) = '',
    @AceptaMovimiento BIT,
    @RequiereCentroCosto BIT,
    @Estado BIT,
    @Orden INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Codigo VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CodigoCuenta)), '');
        DECLARE @Padre VARCHAR(20) = NULLIF(LTRIM(RTRIM(@CodigoCuentaPadre)), '');
        DECLARE @Nivel TINYINT = 1;

        IF @Codigo IS NULL OR NULLIF(LTRIM(RTRIM(@NombreCuenta)), N'') IS NULL
            RAISERROR(N'El codigo y el nombre de la cuenta son obligatorios.', 16, 1);

        IF UPPER(ISNULL(@ColBalance, '')) NOT IN ('S', 'I', 'N', 'F', 'R')
            RAISERROR(N'La columna de balance indicada no es valida.', 16, 1);

        IF UPPER(ISNULL(@IdMoneda, '')) NOT IN ('', 'PEN', 'USD')
            RAISERROR(N'La moneda de la cuenta debe ser PEN, USD o quedar vacia.', 16, 1);

        IF UPPER(ISNULL(@TipoCambio, '')) NOT IN ('', 'C', 'V')
            RAISERROR(N'El tipo de cambio debe ser Compra, Venta o quedar vacio.', 16, 1);

        IF @IdPlanCuentaMaestro IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuentaMaestro WHERE IdPlanCuentaMaestro = @IdPlanCuentaMaestro)
            RAISERROR(N'La cuenta maestra indicada no existe.', 16, 1);

        IF @IdPlanCuentaMaestro IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.CON_PlanCuentaMaestro
               WHERE IdPlanCuentaMaestro = @IdPlanCuentaMaestro
                 AND CodigoCuenta <> @Codigo
           )
            RAISERROR(N'El codigo contable no puede modificarse despues de crear la cuenta maestra.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuentaMaestro
            WHERE CodigoCuenta = @Codigo
              AND (@IdPlanCuentaMaestro IS NULL OR IdPlanCuentaMaestro <> @IdPlanCuentaMaestro)
        )
            RAISERROR(N'Ya existe una cuenta maestra con el codigo indicado.', 16, 1);

        IF @Padre = @Codigo
            RAISERROR(N'Una cuenta no puede ser su propio padre.', 16, 1);

        IF @IdPlanCuentaMaestro IS NOT NULL AND @Padre IS NOT NULL
        BEGIN
            ;WITH Descendientes AS
            (
                SELECT hija.CodigoCuenta
                FROM dbo.CON_PlanCuentaMaestro AS hija
                WHERE hija.CodigoCuentaPadre = @Codigo

                UNION ALL

                SELECT hija.CodigoCuenta
                FROM dbo.CON_PlanCuentaMaestro AS hija
                INNER JOIN Descendientes AS padre
                    ON padre.CodigoCuenta = hija.CodigoCuentaPadre
            )
            SELECT @Nivel = @Nivel
            FROM Descendientes AS descendiente
            WHERE descendiente.CodigoCuenta = @Padre;

            IF @@ROWCOUNT > 0
                RAISERROR(N'La cuenta padre seleccionada pertenece a la propia rama de la cuenta.', 16, 1);
        END;

        IF @Padre IS NOT NULL
        BEGIN
            SELECT @Nivel = CAST(NivelCuenta + 1 AS TINYINT)
            FROM dbo.CON_PlanCuentaMaestro
            WHERE CodigoCuenta = @Padre
              AND Estado = 1;

            IF @@ROWCOUNT = 0
                RAISERROR(N'La cuenta padre no existe o esta inactiva.', 16, 1);

            IF @Nivel > 9
                RAISERROR(N'El nivel calculado excede el maximo permitido.', 16, 1);

            IF EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuentaMaestro
                WHERE CodigoCuenta = @Padre
                  AND AceptaMovimiento = 1
            )
                RAISERROR(N'La cuenta padre no puede aceptar movimiento.', 16, 1);
        END;

        IF @AceptaMovimiento = 1
           AND EXISTS (SELECT 1 FROM dbo.CON_PlanCuentaMaestro WHERE CodigoCuentaPadre = @Codigo)
            RAISERROR(N'Una cuenta con cuentas hijas no puede aceptar movimiento.', 16, 1);

        IF @IdPlanCuentaMaestro IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.CON_PlanCuentaMaestro AS actual
               WHERE actual.IdPlanCuentaMaestro = @IdPlanCuentaMaestro
                 AND ISNULL(actual.CodigoCuentaPadre, '') <> ISNULL(@Padre, '')
           )
           AND EXISTS (SELECT 1 FROM dbo.CON_PlanCuentaMaestro WHERE CodigoCuentaPadre = @Codigo)
            RAISERROR(N'No se puede cambiar la cuenta padre mientras la cuenta tenga cuentas hijas.', 16, 1);

        IF @Estado = 0
           AND
           (
               EXISTS (SELECT 1 FROM dbo.CON_PlanCuentaMaestro WHERE CodigoCuentaPadre = @Codigo AND Estado = 1)
               OR EXISTS (SELECT 1 FROM dbo.CON_CuentaDestinoReglaMaestro WHERE CodigoCuentaOrigen = @Codigo AND Activo = 1)
               OR EXISTS
               (
                   SELECT 1 FROM dbo.CON_CuentaDestinoReglaDetalleMaestro
                   WHERE Activo = 1 AND (CodigoCuentaDestinoCargo = @Codigo OR CodigoCuentaDestinoAbono = @Codigo)
               )
               OR EXISTS (SELECT 1 FROM dbo.CON_TipoImpuesto WHERE CodigoCuenta = @Codigo AND Estado = 1)
               OR EXISTS
               (
                   SELECT 1 FROM dbo.ADM_TipoComprobante
                   WHERE Estado = 1
                     AND (CodigoCuentaVentaSoles = @Codigo OR CodigoCuentaVentaDolares = @Codigo
                          OR CodigoCuentaCompraSoles = @Codigo OR CodigoCuentaCompraDolares = @Codigo)
               )
               OR EXISTS
               (
                   SELECT 1 FROM dbo.ADM_ParametroMaestro
                   WHERE Activo = 1 AND ValorParametro = @Codigo
               )
           )
            RAISERROR(N'No se puede desactivar la cuenta porque tiene hijos activos o configuraciones maestras vigentes.', 16, 1);

        IF @IdPlanCuentaMaestro IS NULL
        BEGIN
            INSERT INTO dbo.CON_PlanCuentaMaestro
            (
                CodigoCuenta, CodigoCuentaPadre, NombreCuenta, NivelCuenta, ColBalance,
                IdMoneda, TipoCambio, AceptaMovimiento, RequiereCentroCosto,
                Estado, Orden, UsuarioRegistro
            )
            VALUES
            (
                @Codigo, @Padre, LTRIM(RTRIM(@NombreCuenta)), @Nivel, UPPER(@ColBalance),
                UPPER(ISNULL(@IdMoneda, '')), UPPER(ISNULL(@TipoCambio, '')), @AceptaMovimiento,
                @RequiereCentroCosto, @Estado, @Orden, @UsuarioRegistro
            );

            SET @IdPlanCuentaMaestro = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.CON_PlanCuentaMaestro
            SET CodigoCuentaPadre = @Padre,
                NombreCuenta = LTRIM(RTRIM(@NombreCuenta)),
                NivelCuenta = @Nivel,
                ColBalance = UPPER(@ColBalance),
                IdMoneda = UPPER(ISNULL(@IdMoneda, '')),
                TipoCambio = UPPER(ISNULL(@TipoCambio, '')),
                AceptaMovimiento = @AceptaMovimiento,
                RequiereCentroCosto = @RequiereCentroCosto,
                Estado = @Estado,
                Orden = @Orden,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdPlanCuentaMaestro = @IdPlanCuentaMaestro;
        END;

        SELECT @IdPlanCuentaMaestro AS IdPlanCuentaMaestro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
