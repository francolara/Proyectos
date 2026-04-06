USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_ConfiguracionClub_Actualizar por ubigeo fiscal y tipo de documento SUNAT centralizado.
CREATE OR ALTER PROCEDURE dbo.Sp_ConfiguracionClub_Actualizar
    @NegocioId INT,
    @NombreComercial NVARCHAR(200),
    @RazonSocial NVARCHAR(200) = NULL,
    @TipoDocumentoFiscal NVARCHAR(20) = NULL,
    @NumeroDocumentoFiscal NVARCHAR(20) = NULL,
    @DireccionFiscal NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @MonedaId INT,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @DireccionFiscalNormalizada NVARCHAR(250);
        DECLARE @CodigoUbigeoNormalizado CHAR(6);

        SET @TipoDocumentoFiscal = NULLIF(UPPER(LTRIM(RTRIM(@TipoDocumentoFiscal))), N'');
        SET @DireccionFiscalNormalizada = NULLIF(LTRIM(RTRIM(@DireccionFiscal)), N'');
        SET @CodigoUbigeoNormalizado = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');

        IF NOT EXISTS (SELECT 1 FROM dbo.Monedas WHERE Id = @MonedaId AND Activo = 1)
            RAISERROR('La moneda seleccionada no es valida.', 16, 1);

        IF @TipoDocumentoFiscal IS NULL
            RAISERROR('El tipo de documento SUNAT es obligatorio.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = @TipoDocumentoFiscal AND t.Activo = 1)
            RAISERROR('El tipo de documento SUNAT no es valido.', 16, 1);

        IF @DireccionFiscalNormalizada IS NULL
            SET @CodigoUbigeoNormalizado = NULL;

        IF @DireccionFiscalNormalizada IS NOT NULL AND @CodigoUbigeoNormalizado IS NULL
            RAISERROR('Cuando se registra direccion fiscal, el distrito es obligatorio.', 16, 1);

        IF @CodigoUbigeoNormalizado IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @CodigoUbigeoNormalizado AND Activo = 1)
            RAISERROR('El codigo de ubigeo no existe.', 16, 1);

        UPDATE n
        SET
            n.NombreComercial = @NombreComercial,
            n.RazonSocial = NULLIF(@RazonSocial, N''),
            n.TipoDocumentoFiscal = @TipoDocumentoFiscal,
            n.NumeroDocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.DireccionFiscal = @DireccionFiscalNormalizada,
            n.CodigoUbigeo = @CodigoUbigeoNormalizado,
            n.DocumentoFiscal = NULLIF(@NumeroDocumentoFiscal, N''),
            n.MonedaId = @MonedaId
        FROM dbo.Negocios n
        WHERE n.Id = @NegocioId
          AND n.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el club para actualizar.', 16, 1);

        DECLARE @EntidadIdAuditoria NVARCHAR(80);
        SET @EntidadIdAuditoria = CONVERT(NVARCHAR(80), @NegocioId);

        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'CONFIGURACION',
            @Accion = N'EDIT',
            @Entidad = N'Negocio',
            @EntidadId = @EntidadIdAuditoria,
            @Usuario = @Usuario,
            @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
