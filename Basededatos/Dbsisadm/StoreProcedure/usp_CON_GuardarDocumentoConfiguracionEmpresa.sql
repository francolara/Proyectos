-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Guarda cuentas contables por tipo de comprobante y empresa.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   20/06/2026
-- Description:   Guarda cuentas de documento separadas para compras y ventas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarDocumentoConfiguracionEmpresa
    @IdEmpresa INT,
    @IdTipoComprobante INT,
    @IdCuentaVentaSoles INT = NULL,
    @IdCuentaVentaDolares INT = NULL,
    @IdCuentaCompraSoles INT = NULL,
    @IdCuentaCompraDolares INT = NULL,
    @Activo BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS (SELECT 1 FROM dbo.ADM_TipoComprobante AS t WHERE t.IdTipoComprobante = @IdTipoComprobante)
        BEGIN
            RAISERROR(N'El documento indicado no existe.', 16, 1);
        END;

        IF @IdCuentaVentaSoles IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuenta AS p WHERE p.IdPlanCuenta = @IdCuentaVentaSoles AND p.IdEmpresa = @IdEmpresa AND p.Estado = 1 AND p.AceptaMovimiento = 1)
        BEGIN
            RAISERROR(N'La cuenta de venta en soles no pertenece a la empresa o no acepta movimiento.', 16, 1);
        END;

        IF @IdCuentaVentaDolares IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuenta AS p WHERE p.IdPlanCuenta = @IdCuentaVentaDolares AND p.IdEmpresa = @IdEmpresa AND p.Estado = 1 AND p.AceptaMovimiento = 1)
        BEGIN
            RAISERROR(N'La cuenta de venta en dolares no pertenece a la empresa o no acepta movimiento.', 16, 1);
        END;

        IF @IdCuentaCompraSoles IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuenta AS p WHERE p.IdPlanCuenta = @IdCuentaCompraSoles AND p.IdEmpresa = @IdEmpresa AND p.Estado = 1 AND p.AceptaMovimiento = 1)
        BEGIN
            RAISERROR(N'La cuenta de compra en soles no pertenece a la empresa o no acepta movimiento.', 16, 1);
        END;

        IF @IdCuentaCompraDolares IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_PlanCuenta AS p WHERE p.IdPlanCuenta = @IdCuentaCompraDolares AND p.IdEmpresa = @IdEmpresa AND p.Estado = 1 AND p.AceptaMovimiento = 1)
        BEGIN
            RAISERROR(N'La cuenta de compra en dolares no pertenece a la empresa o no acepta movimiento.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_DocumentoConfiguracionEmpresa AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.IdTipoComprobante = @IdTipoComprobante
        )
        BEGIN
            UPDATE dbo.CON_DocumentoConfiguracionEmpresa
            SET IdCuentaVentaSoles = @IdCuentaVentaSoles,
                IdCuentaVentaDolares = @IdCuentaVentaDolares,
                IdCuentaCompraSoles = @IdCuentaCompraSoles,
                IdCuentaCompraDolares = @IdCuentaCompraDolares,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdTipoComprobante = @IdTipoComprobante;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.CON_DocumentoConfiguracionEmpresa
            (
                IdEmpresa,
                IdTipoComprobante,
                IdCuentaVentaSoles,
                IdCuentaVentaDolares,
                IdCuentaCompraSoles,
                IdCuentaCompraDolares,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @IdTipoComprobante,
                @IdCuentaVentaSoles,
                @IdCuentaVentaDolares,
                @IdCuentaCompraSoles,
                @IdCuentaCompraDolares,
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
