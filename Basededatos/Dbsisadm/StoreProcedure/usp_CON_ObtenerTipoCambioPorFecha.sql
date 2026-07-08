-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Obtiene el tipo de cambio vigente por cuenta administradora, fecha y moneda para autocompletar registros contables.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Permite consultar el tipo de cambio operativo por fecha y moneda desde compras, ventas, asientos y caja y bancos.

CREATE OR ALTER PROCEDURE dbo.usp_CON_ObtenerTipoCambioPorFecha
    @IdCuentaAdministradora INT,
    @Fecha DATE,
    @IdMoneda VARCHAR(3)
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT TOP (1)
            tc.IdTipoCambio,
            tc.IdCuentaAdministradora,
            tc.Fecha,
            tc.IdMoneda,
            tc.Compra,
            tc.Venta,
            tc.CompraSBS,
            tc.VentaSBS,
            tc.Fuente,
            tc.UsuarioRegistro,
            tc.Estado
        FROM dbo.CON_TipoCambio AS tc
        WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
          AND tc.Fecha = @Fecha
          AND tc.IdMoneda = UPPER(LTRIM(RTRIM(@IdMoneda)))
          AND tc.Estado = 1
        ORDER BY
            tc.IdTipoCambio DESC;

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
