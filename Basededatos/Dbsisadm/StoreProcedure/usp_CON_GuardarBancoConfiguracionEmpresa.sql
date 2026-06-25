-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Inserta o actualiza una cuenta corriente bancaria por empresa validando banco, numero, titular, moneda y cuenta contable.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarBancoConfiguracionEmpresa
    @IdBancoConfiguracionEmpresa INT = NULL,
    @IdEmpresa INT,
    @IdBanco INT,
    @NroCuentaCorriente VARCHAR(50),
    @Titular VARCHAR(200),
    @IdMoneda INT,
    @IdPlanCuenta INT,
    @Activo BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Bancos AS b
            WHERE b.IdBanco = @IdBanco
              AND b.Estado = 1
        )
        BEGIN
            RAISERROR(N'El banco seleccionado no existe o esta inactivo.', 16, 1);
        END;

        IF NULLIF(LTRIM(RTRIM(@Titular)), '') IS NULL
        BEGIN
            RAISERROR(N'Ingrese el titular de la cuenta corriente.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Moneda AS m
            WHERE m.IdMoneda = @IdMoneda
              AND m.Estado = 1
        )
        BEGIN
            RAISERROR(N'La moneda seleccionada no existe o esta inactiva.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS p
            WHERE p.IdPlanCuenta = @IdPlanCuenta
              AND p.IdEmpresa = @IdEmpresa
              AND p.Estado = 1
              AND p.AceptaMovimiento = 1
        )
        BEGIN
            RAISERROR(N'La cuenta contable asociada no existe, esta inactiva o no acepta movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_BancosConfiguracionEmpresa AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.NroCuentaCorriente = @NroCuentaCorriente
              AND (@IdBancoConfiguracionEmpresa IS NULL OR c.IdBancoConfiguracionEmpresa <> @IdBancoConfiguracionEmpresa)
        )
        BEGIN
            RAISERROR(N'Ya existe una cuenta corriente con el mismo numero para la empresa activa.', 16, 1);
        END;

        IF @IdBancoConfiguracionEmpresa IS NULL
        BEGIN
            INSERT INTO dbo.CON_BancosConfiguracionEmpresa
            (
                IdEmpresa,
                IdBanco,
                NroCuentaCorriente,
                Titular,
                IdMoneda,
                IdPlanCuenta,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdBanco,
                @NroCuentaCorriente,
                LTRIM(RTRIM(@Titular)),
                @IdMoneda,
                @IdPlanCuenta,
                @Activo,
                @UsuarioRegistro
            );

            SET @IdBancoConfiguracionEmpresa = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_BancosConfiguracionEmpresa AS c
                WHERE c.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa
                  AND c.IdEmpresa = @IdEmpresa
            )
            BEGIN
                RAISERROR(N'La cuenta corriente a actualizar no existe en la empresa activa.', 16, 1);
            END;

            UPDATE dbo.CON_BancosConfiguracionEmpresa
            SET IdBanco = @IdBanco,
                NroCuentaCorriente = @NroCuentaCorriente,
                Titular = LTRIM(RTRIM(@Titular)),
                IdMoneda = @IdMoneda,
                IdPlanCuenta = @IdPlanCuenta,
                Activo = @Activo
            WHERE IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa
              AND IdEmpresa = @IdEmpresa;
        END;

        SELECT
            c.IdBancoConfiguracionEmpresa,
            c.IdEmpresa,
            c.IdBanco,
            b.Codigo AS CodigoBanco,
            b.Nombre AS NombreBanco,
            c.NroCuentaCorriente,
            c.Titular,
            c.IdMoneda,
            m.CodigoMoneda,
            m.NombreMoneda,
            c.IdPlanCuenta,
            p.CodigoCuenta,
            p.NombreCuenta,
            c.Activo,
            c.FechaRegistro,
            c.UsuarioRegistro
        FROM dbo.CON_BancosConfiguracionEmpresa AS c
        INNER JOIN dbo.CON_Bancos AS b
            ON b.IdBanco = c.IdBanco
        INNER JOIN dbo.CON_PlanCuenta AS p
            ON p.IdPlanCuenta = c.IdPlanCuenta
        INNER JOIN dbo.ADM_Moneda AS m
            ON m.IdMoneda = c.IdMoneda
        WHERE c.IdBancoConfiguracionEmpresa = @IdBancoConfiguracionEmpresa;

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
