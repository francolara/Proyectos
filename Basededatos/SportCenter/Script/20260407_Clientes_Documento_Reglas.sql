USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:        FRANCO LARA
-- Create date:   07/04/2026
-- Description:   Reglas de tipo/numero de documento en clientes (max 11 digitos y tipo no domiciliado sin numero).
-- Firma:         Codex - 07/04/2026
-- =============================================

CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Crear
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
    @Nombres NVARCHAR(120) = NULL,
    @Apellidos NVARCHAR(120) = NULL,
    @NombreEquipo NVARCHAR(120) = NULL,
    @TipoDocumento NVARCHAR(20),
    @NumeroDocumento NVARCHAR(20),
    @Telefono NVARCHAR(20) = NULL,
    @Correo NVARCHAR(200) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @Activo BIT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @NumeroDocumentoNormalizado NVARCHAR(20);
        DECLARE @NombreEquipoNormalizado NVARCHAR(120);
        DECLARE @DireccionFiscalNormalizada NVARCHAR(250);
        DECLARE @CodigoUbigeoNormalizado CHAR(6);
        DECLARE @NombresNormalizado NVARCHAR(120);
        DECLARE @ApellidosNormalizado NVARCHAR(120);
        DECLARE @NombresORazonSocialNormalizado NVARCHAR(200);

        SET @TipoDocumento = UPPER(LTRIM(RTRIM(@TipoDocumento)));
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NombreEquipoNormalizado = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        SET @DireccionFiscalNormalizada = NULLIF(LTRIM(RTRIM(@DireccionFiscal)), N'');
        SET @CodigoUbigeoNormalizado = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');
        SET @NombresNormalizado = NULLIF(LTRIM(RTRIM(@Nombres)), N'');
        SET @ApellidosNormalizado = NULLIF(LTRIM(RTRIM(@Apellidos)), N'');
        SET @NombresORazonSocialNormalizado = NULLIF(LTRIM(RTRIM(@NombresORazonSocial)), N'');

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = @TipoDocumento AND t.Activo = 1)
            RAISERROR('El tipo de documento SUNAT no es valido.', 16, 1);

        IF @TipoDocumento = N'0'
        BEGIN
            SET @NumeroDocumentoNormalizado = NULL;
            SET @NumeroDocumento = N'';
        END
        ELSE
        BEGIN
            IF @NumeroDocumentoNormalizado IS NULL
                RAISERROR('Ingresa el numero de documento.', 16, 1);

            IF LEN(@NumeroDocumentoNormalizado) > 11
                RAISERROR('El numero de documento permite como maximo 11 digitos.', 16, 1);

            IF @NumeroDocumentoNormalizado LIKE N'%[^0-9]%'
                RAISERROR('El numero de documento solo permite digitos.', 16, 1);
        END;

        IF @TipoDocumento = N'6'
        BEGIN
            IF @NombresORazonSocialNormalizado IS NULL
                RAISERROR('Ingresa la razon social para tipo de documento RUC.', 16, 1);

            SET @NombresNormalizado = NULL;
            SET @ApellidosNormalizado = NULL;
        END
        ELSE
        BEGIN
            IF @NombresNormalizado IS NULL
                RAISERROR('Ingresa los nombres del cliente.', 16, 1);

            IF @ApellidosNormalizado IS NULL
                RAISERROR('Ingresa los apellidos del cliente.', 16, 1);

            SET @NombresORazonSocialNormalizado = LEFT(LTRIM(RTRIM(CONCAT(@NombresNormalizado, N' ', @ApellidosNormalizado))), 200);
        END;

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               WHERE c.NegocioId = @NegocioId
                 AND c.Activo = 1
                 AND LTRIM(RTRIM(c.NumeroDocumento)) = @NumeroDocumentoNormalizado
           )
            RAISERROR('Cliente ya se encuentra registrado.', 16, 1);

        IF @DireccionFiscalNormalizada IS NULL
            SET @CodigoUbigeoNormalizado = NULL;

        IF @DireccionFiscalNormalizada IS NOT NULL AND @CodigoUbigeoNormalizado IS NULL
            RAISERROR('Cuando se registra direccion fiscal, el distrito es obligatorio.', 16, 1);

        IF @CodigoUbigeoNormalizado IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeoNormalizado AND Activo = 1)
            RAISERROR('El codigo de ubigeo no existe.', 16, 1);

        BEGIN TRANSACTION;

        INSERT INTO dbo.Clientes
        (
            NegocioId, NombresORazonSocial, Nombres, Apellidos, NombreEquipo, TipoDocumento, NumeroDocumento, Telefono,
            Correo, DireccionFiscal, CodigoUbigeo, Activo, FechaCreacion, UsuarioCreacion
        )
        VALUES
        (
            @NegocioId, @NombresORazonSocialNormalizado, @NombresNormalizado, @ApellidosNormalizado, @NombreEquipoNormalizado, @TipoDocumento, @NumeroDocumento, @Telefono,
            @Correo, @DireccionFiscalNormalizada, @CodigoUbigeoNormalizado, @Activo, SYSUTCDATETIME(), @Usuario
        );

        DECLARE @Id INT;
        SET @Id = SCOPE_IDENTITY();

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'CREATE', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;

        COMMIT TRANSACTION;

        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
