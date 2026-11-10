-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Registra abonos idempotentes con frecuencia y periodo de cobertura.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/11/2026
-- Description:   Admite frecuencia sin fechas de cobertura.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PagoProveedor_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @CostoProveedorId BIGINT,
    @Importe DECIMAL(20,4),
    @FechaPago DATE,
    @Referencia NVARCHAR(100),
    @Frecuencia VARCHAR(10) = NULL,
    @PeriodoDesde DATE = NULL,
    @PeriodoHasta DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PAGOS_PROVEEDOR.EDITAR') = 0
            THROW 51000, N'No tiene permiso para registrar pagos a proveedores.', 1;

        SET @Referencia = TRIM(@Referencia);
        IF (@Frecuencia IS NOT NULL AND @Frecuencia NOT IN ('SEMANAL', 'QUINCENAL', 'MENSUAL'))
           OR ((@PeriodoDesde IS NOT NULL OR @PeriodoHasta IS NOT NULL)
               AND (@Frecuencia IS NULL OR @PeriodoDesde IS NULL OR @PeriodoHasta IS NULL OR @PeriodoHasta < @PeriodoDesde))
            THROW 51000, N'Revise la frecuencia y las fechas del periodo.', 1;
        IF @OperacionId IS NULL OR @FechaPago IS NULL OR NULLIF(@Referencia, N'') IS NULL
           OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
            THROW 51000, N'Indique fecha, referencia e importe positivo con hasta dos decimales.', 1;

        BEGIN TRANSACTION;

        -- Serializa los abonos de un mismo costo hasta confirmar el pago y su auditoria.
        DECLARE @Costo DECIMAL(18,2);
        SELECT @Costo = costo.Importe
        FROM dbo.CostoProveedorPresupuesto AS costo WITH (UPDLOCK, HOLDLOCK)
        WHERE costo.CostoProveedorId = @CostoProveedorId;

        IF @Costo IS NULL
            THROW 51000, N'El costo del proveedor no existe.', 1;

        DECLARE @PagoId BIGINT;
        SELECT @PagoId = pago.PagoProveedorId
        FROM dbo.PagoProveedorHYT AS pago
        WHERE pago.OperacionId = @OperacionId;

        IF @PagoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.PagoProveedorHYT AS pago
                WHERE pago.PagoProveedorId = @PagoId AND pago.UsuarioRegistro = @UsuarioId
                    AND pago.CostoProveedorId = @CostoProveedorId AND pago.Importe = @Importe
                    AND pago.FechaPago = @FechaPago AND pago.Referencia = @Referencia
                    AND pago.Frecuencia IS NOT DISTINCT FROM @Frecuencia
                    AND pago.PeriodoDesde IS NOT DISTINCT FROM @PeriodoDesde
                    AND pago.PeriodoHasta IS NOT DISTINCT FROM @PeriodoHasta)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            DECLARE @Pagado DECIMAL(18,2);
            SELECT @Pagado = COALESCE(SUM(pago.Importe), 0)
            FROM dbo.PagoProveedorHYT AS pago
            WHERE pago.CostoProveedorId = @CostoProveedorId;

            IF @Importe > @Costo - @Pagado
                THROW 51000, N'El importe excede el saldo pendiente.', 1;

            INSERT INTO dbo.PagoProveedorHYT
                (CostoProveedorId, Importe, FechaPago, Referencia, OperacionId, UsuarioRegistro,
                 Frecuencia, PeriodoDesde, PeriodoHasta)
            VALUES (@CostoProveedorId, @Importe, @FechaPago, @Referencia, @OperacionId, @UsuarioId,
                    @Frecuencia, @PeriodoDesde, @PeriodoHasta);

            SET @PagoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PagoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'PagoProveedorHYT', @RegistroId, N'REGISTRAR';
        END;

        COMMIT TRANSACTION;
        SELECT @PagoId AS PagoProveedorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
