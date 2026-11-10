-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_CostoProveedor_Registrar para Planillas.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_CostoProveedor_Registrar
    @UsuarioId NVARCHAR(128),
    @OperacionId UNIQUEIDENTIFIER,
    @PresupuestoId BIGINT,
    @ProveedorId BIGINT,
    @Descripcion NVARCHAR(300),
    @Importe DECIMAL(20,4),
    @MonedaCodigo NVARCHAR(3),
    @DocumentoSustento NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'COSTOS_PROVEEDOR.EDITAR') = 0
           OR dbo.fn_HYTPermiso(@UsuarioId, N'PRESUPUESTOS.CONSULTAR') = 0
            THROW 51000, N'No tiene permiso para registrar costos del presupuesto.', 1;

        SET @Descripcion = TRIM(@Descripcion);
        SET @DocumentoSustento = NULLIF(TRIM(@DocumentoSustento), N'');
        IF @OperacionId IS NULL OR NULLIF(@Descripcion, N'') IS NULL
           OR @Importe IS NULL OR @Importe <= 0 OR @Importe <> ROUND(@Importe, 2)
            THROW 51000, N'Indique una descripcion y un importe positivo con hasta dos decimales.', 1;

        BEGIN TRANSACTION;

        DECLARE @CostoId BIGINT;
        SELECT @CostoId = costo.CostoProveedorId
        FROM dbo.CostoProveedorPresupuesto AS costo WITH (UPDLOCK, HOLDLOCK)
        WHERE costo.OperacionId = @OperacionId;

        IF @CostoId IS NOT NULL
        BEGIN
            IF NOT EXISTS (
                SELECT 1 FROM dbo.CostoProveedorPresupuesto AS costo
                WHERE costo.CostoProveedorId = @CostoId AND costo.UsuarioRegistro = @UsuarioId
                    AND costo.PresupuestoId = @PresupuestoId AND costo.ProveedorId = @ProveedorId
                    AND costo.Descripcion = @Descripcion AND costo.Importe = @Importe
                    AND costo.MonedaCodigo = @MonedaCodigo
                    AND costo.DocumentoSustento IS NOT DISTINCT FROM @DocumentoSustento)
                THROW 51000, N'La operacion ya se utilizo con otros datos.', 1;
        END
        ELSE
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.ProveedorHYT AS proveedor WITH (HOLDLOCK)
                           WHERE proveedor.ProveedorId = @ProveedorId AND proveedor.Activo = 1)
                THROW 51000, N'Seleccione un proveedor activo.', 1;

            INSERT INTO dbo.CostoProveedorPresupuesto
                (PresupuestoId, ProveedorId, Descripcion, Importe, MonedaCodigo, DocumentoSustento, OperacionId, UsuarioRegistro)
            VALUES (@PresupuestoId, @ProveedorId, @Descripcion, @Importe, @MonedaCodigo,
                    @DocumentoSustento, @OperacionId, @UsuarioId);

            SET @CostoId = SCOPE_IDENTITY();
            DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @CostoId);
            EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @OperacionId, N'CostoProveedorPresupuesto', @RegistroId, N'REGISTRAR';
        END;

        COMMIT TRANSACTION;
        SELECT @CostoId AS CostoProveedorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
