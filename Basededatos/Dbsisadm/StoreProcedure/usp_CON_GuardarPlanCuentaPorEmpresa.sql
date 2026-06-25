-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Inserta o actualiza una cuenta contable por empresa validando codigo, padre y nivel.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Reemplaza NaturalezaSaldo por ColBalance, agrega IdMoneda/TipoCambio y valida longitud por parametros de grado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarPlanCuentaPorEmpresa
    @IdPlanCuenta INT = NULL,
    @IdEmpresa INT,
    @IdPlanCuentaPadre INT = NULL,
    @CodigoCuenta VARCHAR(20),
    @NombreCuenta NVARCHAR(200),
    @ColBalance CHAR(1),
    @IdMoneda VARCHAR(3) = '',
    @TipoCambio CHAR(1) = '',
    @AceptaMovimiento BIT,
    @RequiereCentroCosto BIT,
    @Estado BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @NivelCuenta TINYINT = 1
        DECLARE @CodigoCuentaTrabajo VARCHAR(20) = LTRIM(RTRIM(@CodigoCuenta))
        DECLARE @CodigoCuentaPadre VARCHAR(20) = NULL
        DECLARE @GradoMaximo TINYINT
        DECLARE @LongitudEsperada INT = 0
        DECLARE @NivelIterador TINYINT = 1
        DECLARE @CodigoParametroLongitud VARCHAR(100)
        DECLARE @LongitudNivel INT

        IF @ColBalance NOT IN ('S', 'I', 'N', 'F', 'R')
        BEGIN
            RAISERROR(N'La columna de balance es invalida.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF ISNULL(@IdMoneda, '') NOT IN ('', 'PEN', 'USD')
        BEGIN
            RAISERROR(N'La moneda de la cuenta es invalida.', 16, 1);
        END;

        IF ISNULL(@TipoCambio, '') NOT IN ('', 'V', 'C')
        BEGIN
            RAISERROR(N'El tipo de cambio de la cuenta es invalido.', 16, 1);
        END;

        IF @IdPlanCuentaPadre IS NOT NULL
        BEGIN
            SELECT
                @NivelCuenta = pc.NivelCuenta + 1,
                @CodigoCuentaPadre = pc.CodigoCuenta
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdPlanCuenta = @IdPlanCuentaPadre
              AND pc.IdEmpresa = @IdEmpresa;

            IF @NivelCuenta = 1
            BEGIN
                RAISERROR(N'La cuenta padre no pertenece a la empresa activa.', 16, 1);
            END;
        END;

        SELECT
            @GradoMaximo = TRY_CONVERT(TINYINT, pe.ValorParametro)
        FROM dbo.ADM_ParametroEmpresa AS pe
        WHERE pe.IdEmpresa = @IdEmpresa
          AND pe.TipoParametro = 'CONTABLE'
          AND pe.CodigoParametro = 'GRADO_MAXIMO'
          AND pe.Activo = 1;

        IF @GradoMaximo IS NULL OR @GradoMaximo <= 0
        BEGIN
            RAISERROR(N'Configure el parametro contable GRADO_MAXIMO para la empresa activa.', 16, 1);
        END;

        IF @NivelCuenta > @GradoMaximo
        BEGIN
            RAISERROR(N'El nivel de la cuenta supera el grado maximo configurado para la empresa.', 16, 1);
        END;

        WHILE @NivelIterador <= @NivelCuenta
        BEGIN
            SET @CodigoParametroLongitud = CONCAT('GRADO', @NivelIterador, '_LONG');
            SET @LongitudNivel = NULL;

            SELECT
                @LongitudNivel = TRY_CONVERT(INT, pe.ValorParametro)
            FROM dbo.ADM_ParametroEmpresa AS pe
            WHERE pe.IdEmpresa = @IdEmpresa
              AND pe.TipoParametro = 'CONTABLE'
              AND pe.CodigoParametro = @CodigoParametroLongitud
              AND pe.Activo = 1;

            IF @LongitudNivel IS NULL OR @LongitudNivel <= 0
            BEGIN
                RAISERROR(N'Configure la longitud del grado correspondiente en parametros contables.', 16, 1);
            END;

            SET @LongitudEsperada += @LongitudNivel;
            SET @NivelIterador += 1;
        END;

        IF LEN(@CodigoCuentaTrabajo) <> @LongitudEsperada
        BEGIN
            RAISERROR(N'El codigo de cuenta no cumple la longitud configurada para el nivel calculado.', 16, 1);
        END;

        IF @CodigoCuentaPadre IS NOT NULL
           AND LEFT(@CodigoCuentaTrabajo, LEN(@CodigoCuentaPadre)) <> @CodigoCuentaPadre
        BEGIN
            RAISERROR(N'El codigo de cuenta debe iniciar con el codigo de la cuenta padre seleccionada.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.CodigoCuenta = @CodigoCuentaTrabajo
              AND (@IdPlanCuenta IS NULL OR pc.IdPlanCuenta <> @IdPlanCuenta)
        )
        BEGIN
            RAISERROR(N'Ya existe una cuenta con el mismo codigo para la empresa activa.', 16, 1);
        END;

        IF @IdPlanCuenta IS NULL
        BEGIN
            INSERT INTO dbo.CON_PlanCuenta
            (
                IdEmpresa,
                IdPlanCuentaPadre,
                CodigoCuenta,
                NombreCuenta,
                NivelCuenta,
                ColBalance,
                IdMoneda,
                TipoCambio,
                AceptaMovimiento,
                RequiereCentroCosto,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdPlanCuentaPadre,
                @CodigoCuentaTrabajo,
                @NombreCuenta,
                @NivelCuenta,
                @ColBalance,
                ISNULL(@IdMoneda, ''),
                ISNULL(@TipoCambio, ''),
                @AceptaMovimiento,
                @RequiereCentroCosto,
                @Estado,
                @UsuarioRegistro
            );

            SET @IdPlanCuenta = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_PlanCuenta AS pc
                WHERE pc.IdPlanCuenta = @IdPlanCuenta
                  AND pc.IdEmpresa = @IdEmpresa
            )
            BEGIN
                RAISERROR(N'La cuenta a actualizar no existe en la empresa activa.', 16, 1);
            END;

            IF @IdPlanCuentaPadre = @IdPlanCuenta
            BEGIN
                RAISERROR(N'La cuenta no puede ser su propio padre.', 16, 1);
            END;

            UPDATE dbo.CON_PlanCuenta
            SET IdPlanCuentaPadre = @IdPlanCuentaPadre,
                CodigoCuenta = @CodigoCuentaTrabajo,
                NombreCuenta = @NombreCuenta,
                NivelCuenta = @NivelCuenta,
                ColBalance = @ColBalance,
                IdMoneda = ISNULL(@IdMoneda, ''),
                TipoCambio = ISNULL(@TipoCambio, ''),
                AceptaMovimiento = @AceptaMovimiento,
                RequiereCentroCosto = @RequiereCentroCosto,
                Estado = @Estado,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdPlanCuenta = @IdPlanCuenta
              AND IdEmpresa = @IdEmpresa;
        END;

        SELECT
            pc.IdPlanCuenta,
            pc.IdEmpresa,
            pc.IdPlanCuentaPadre,
            pc.CodigoCuenta,
            pc.NombreCuenta,
            pc.NivelCuenta,
            pc.ColBalance,
            pc.IdMoneda,
            pc.TipoCambio,
            pc.AceptaMovimiento,
            pc.RequiereCentroCosto,
            pc.Estado
        FROM dbo.CON_PlanCuenta AS pc
        WHERE pc.IdPlanCuenta = @IdPlanCuenta;

    END TRY

    BEGIN CATCH

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
