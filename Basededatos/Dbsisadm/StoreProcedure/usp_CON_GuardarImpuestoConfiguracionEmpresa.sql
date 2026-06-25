-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Guarda cuentas contables por tipo de impuesto, modulo y empresa.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Guarda configuracion unica de cuenta contable por impuesto y empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarImpuestoConfiguracionEmpresa
    @IdEmpresa INT,
    @IdTipoImpuesto INT,
    @IdPlanCuenta INT = NULL,
    @Activo BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS (SELECT 1 FROM dbo.CON_TipoImpuesto AS i WHERE i.IdTipoImpuesto = @IdTipoImpuesto AND i.Estado = 1)
        BEGIN
            RAISERROR(N'El impuesto indicado no existe o no esta activo.', 16, 1);
        END;

        IF @IdPlanCuenta IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuenta AS p WHERE p.IdPlanCuenta = @IdPlanCuenta AND p.IdEmpresa = @IdEmpresa AND p.Estado = 1 AND p.AceptaMovimiento = 1)
        BEGIN
            RAISERROR(N'La cuenta contable no pertenece a la empresa o no acepta movimiento.', 16, 1);
        END;

        IF EXISTS (SELECT 1 FROM dbo.CON_TipoImpuestoConfiguracionEmpresa AS c WHERE c.IdEmpresa = @IdEmpresa AND c.IdTipoImpuesto = @IdTipoImpuesto)
        BEGIN
            UPDATE dbo.CON_TipoImpuestoConfiguracionEmpresa
            SET IdPlanCuenta = @IdPlanCuenta,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdTipoImpuesto = @IdTipoImpuesto;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.CON_TipoImpuestoConfiguracionEmpresa
            (
                IdEmpresa,
                IdTipoImpuesto,
                IdPlanCuenta,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdTipoImpuesto,
                @IdPlanCuenta,
                @Activo,
                @UsuarioRegistro
            );
        END;

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
