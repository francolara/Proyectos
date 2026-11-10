-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Implementa spu_ProveedorHYT_Guardar para Planillas. Incluye el CCI con permiso bancario.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ProveedorHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Ruc VARCHAR(11),
    @RazonSocial NVARCHAR(180),
    @Activo BIT,
    @ProveedorId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @ActualizarCuenta BIT = 0,
    @Banco NVARCHAR(80) = NULL,
    @NumeroCuenta NVARCHAR(34) = NULL,
    @CuentaInterbancaria VARCHAR(100) = NULL,
    @ActualizarCCI BIT = 0
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar proveedores.', 1;

        SET @RazonSocial = NULLIF(TRIM(@RazonSocial), N'');
        IF @Ruc IS NULL OR LEN(@Ruc) <> 11 OR @Ruc LIKE '%[^0-9]%'
           OR @RazonSocial IS NULL OR @Activo IS NULL
            THROW 51000, N'Indique un RUC de once digitos y la razon social.', 1;

        SET @Banco = NULLIF(TRIM(@Banco), N'');
        SET @NumeroCuenta = NULLIF(TRIM(@NumeroCuenta), N'');
        SET @CuentaInterbancaria = NULLIF(TRIM(@CuentaInterbancaria), '');
        IF @ActualizarCCI = 1 AND @ActualizarCuenta <> 1
            THROW 51000, N'El CCI requiere autorizacion para actualizar la cuenta bancaria.', 1;

        IF @ActualizarCuenta = 1
        BEGIN
            IF dbo.fn_HYTPermiso(@UsuarioId, N'PROVEEDORES.CUENTA') = 0
                THROW 51000, N'No tiene permiso para cambiar datos bancarios.', 1;

            IF (@Banco IS NULL AND @NumeroCuenta IS NOT NULL)
               OR (@Banco IS NOT NULL AND @NumeroCuenta IS NULL)
                THROW 51000, N'Indique banco y numero de cuenta juntos.', 1;

            IF @ActualizarCCI = 1 AND @CuentaInterbancaria IS NOT NULL
                AND (@Banco IS NULL OR DATALENGTH(@CuentaInterbancaria) <> 20
                     OR @CuentaInterbancaria COLLATE Latin1_General_100_BIN2 LIKE '%[^0-9]%')
                THROW 51000, N'Indique un banco y un CCI de veinte digitos.', 1;
        END;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @ProveedorId IS NULL
        BEGIN
            INSERT INTO dbo.ProveedorHYT (Ruc, RazonSocial, Banco, NumeroCuenta, CuentaInterbancaria, Activo)
            VALUES (@Ruc, @RazonSocial, CASE WHEN @ActualizarCuenta = 1 THEN @Banco END,
                    CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta END,
                    CASE WHEN @ActualizarCuenta = 1 AND @ActualizarCCI = 1 THEN @CuentaInterbancaria END, @Activo);
            SET @ProveedorId = SCOPE_IDENTITY();
            SET @Accion = N'CREAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.ProveedorHYT
            SET Ruc = @Ruc, RazonSocial = @RazonSocial, Activo = @Activo,
                Banco = CASE WHEN @ActualizarCuenta = 1 THEN @Banco ELSE Banco END,
                CuentaInterbancaria = CASE WHEN @ActualizarCuenta = 1 AND @Banco IS NULL THEN NULL
                    WHEN @ActualizarCuenta = 1 AND @ActualizarCCI = 1 THEN @CuentaInterbancaria ELSE CuentaInterbancaria END,
                NumeroCuenta = CASE WHEN @ActualizarCuenta = 1 THEN @NumeroCuenta ELSE NumeroCuenta END
            WHERE ProveedorId = @ProveedorId AND VersionFila = @VersionFila;

            IF @@ROWCOUNT = 0
                THROW 51000, N'El proveedor cambio o ya no existe. Vuelva a abrir su ficha.', 1;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @ProveedorId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Ruc AS Ruc, @RazonSocial AS RazonSocial, @Activo AS Activo,
                   @ActualizarCuenta AS ActualizoCuenta, @ActualizarCCI AS ActualizoCCI
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'ProveedorHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;

        SELECT @ProveedorId AS ProveedorId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
