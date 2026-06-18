-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Inserta o actualiza una cuenta contable por empresa validando codigo, padre y nivel.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarPlanCuentaPorEmpresa
    @IdPlanCuenta INT = NULL,
    @IdEmpresa INT,
    @IdPlanCuentaPadre INT = NULL,
    @CodigoCuenta VARCHAR(20),
    @NombreCuenta NVARCHAR(200),
    @NaturalezaSaldo CHAR(1),
    @AceptaMovimiento BIT,
    @RequiereCentroCosto BIT,
    @Estado BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @NivelCuenta TINYINT = 1

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF @NaturalezaSaldo NOT IN ('D', 'H')
        BEGIN
            RAISERROR(N'La naturaleza de saldo es invalida.', 16, 1);
        END;

        IF @IdPlanCuentaPadre IS NOT NULL
        BEGIN
            SELECT
                @NivelCuenta = pc.NivelCuenta + 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdPlanCuenta = @IdPlanCuentaPadre
              AND pc.IdEmpresa = @IdEmpresa;

            IF @NivelCuenta = 1
            BEGIN
                RAISERROR(N'La cuenta padre no pertenece a la empresa activa.', 16, 1);
            END;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
              AND pc.CodigoCuenta = @CodigoCuenta
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
                NaturalezaSaldo,
                AceptaMovimiento,
                RequiereCentroCosto,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdPlanCuentaPadre,
                @CodigoCuenta,
                @NombreCuenta,
                @NivelCuenta,
                @NaturalezaSaldo,
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
                CodigoCuenta = @CodigoCuenta,
                NombreCuenta = @NombreCuenta,
                NivelCuenta = @NivelCuenta,
                NaturalezaSaldo = @NaturalezaSaldo,
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
            pc.NaturalezaSaldo,
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
