-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Consulta proveedor y devuelve el CCI solo en fichas autorizadas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ProveedorHYT_Listar
    @UsuarioId NVARCHAR(128),
    @ProveedorId BIGINT = NULL,
    @Buscar NVARCHAR(180) = NULL,
    @SoloActivos BIT = 0,
    @Pagina INT = 1
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para consultar proveedores.', 1;

        IF @Pagina IS NULL OR @Pagina < 1
            THROW 51000, N'Pagina no valida.', 1;

        DECLARE @VerCuenta BIT = dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA');
        SET @Buscar = NULLIF(TRIM(@Buscar), N'');

        SELECT proveedor.ProveedorId, proveedor.Ruc, proveedor.RazonSocial, proveedor.Banco,
               CASE WHEN proveedor.NumeroCuenta IS NOT NULL THEN
                    CONCAT(N'****', RIGHT(proveedor.NumeroCuenta, 4)) END AS CuentaEnmascarada,
               CASE WHEN @ProveedorId IS NOT NULL AND @VerCuenta = 1
                    THEN proveedor.NumeroCuenta END AS NumeroCuenta,
               CASE WHEN @ProveedorId IS NOT NULL AND @VerCuenta = 1
                    THEN proveedor.CuentaInterbancaria END AS CuentaInterbancaria,
               proveedor.Activo, proveedor.VersionFila, COUNT_BIG(1) OVER () AS Total
        FROM dbo.ProveedorHYT AS proveedor
        WHERE (@ProveedorId IS NULL OR proveedor.ProveedorId = @ProveedorId)
          AND (@SoloActivos = 0 OR proveedor.Activo = 1)
          AND (@Buscar IS NULL OR CHARINDEX(@Buscar, proveedor.RazonSocial) > 0
               OR CHARINDEX(@Buscar, proveedor.Ruc) > 0)
        ORDER BY proveedor.RazonSocial, proveedor.ProveedorId
        OFFSET (CONVERT(BIGINT, @Pagina) - 1) * 20 ROWS FETCH NEXT 20 ROWS ONLY;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
