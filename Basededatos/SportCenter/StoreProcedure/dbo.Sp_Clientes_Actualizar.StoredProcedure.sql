USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 04/04/2026 | Actualizacion individual de Sp_Clientes_Actualizar por ubigeo fiscal y tipo de documento SUNAT centralizado.
CREATE OR ALTER PROCEDURE dbo.Sp_Clientes_Actualizar
    @Id INT,
    @NegocioId INT,
    @NombresORazonSocial NVARCHAR(200),
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

        SET @TipoDocumento = UPPER(LTRIM(RTRIM(@TipoDocumento)));
        SET @NumeroDocumentoNormalizado = NULLIF(LTRIM(RTRIM(@NumeroDocumento)), N'');
        SET @NombreEquipoNormalizado = NULLIF(LTRIM(RTRIM(@NombreEquipo)), N'');
        SET @DireccionFiscalNormalizada = NULLIF(LTRIM(RTRIM(@DireccionFiscal)), N'');
        SET @CodigoUbigeoNormalizado = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @NumeroDocumento = COALESCE(@NumeroDocumentoNormalizado, N'');

        IF NOT EXISTS (SELECT 1 FROM dbo.TiposDocumentoIdentidadSunat t WHERE t.CodigoSunat = @TipoDocumento AND t.Activo = 1)
            RAISERROR('El tipo de documento SUNAT no es valido.', 16, 1);

        IF @NumeroDocumentoNormalizado IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.Clientes c
               INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
               WHERE nc.NegocioId = @NegocioId
                 AND nc.Activo = 1
                 AND c.Activo = 1
                 AND c.Id <> @Id
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

        UPDATE c
        SET
            c.NombresORazonSocial = @NombresORazonSocial,
            c.NombreEquipo = @NombreEquipoNormalizado,
            c.TipoDocumento = @TipoDocumento,
            c.NumeroDocumento = @NumeroDocumento,
            c.Telefono = @Telefono,
            c.Correo = @Correo,
            c.DireccionFiscal = @DireccionFiscalNormalizada,
            c.CodigoUbigeo = @CodigoUbigeoNormalizado,
            c.Activo = @Activo,
            c.FechaActualizacion = SYSUTCDATETIME(),
            c.UsuarioActualizacion = @Usuario
        FROM dbo.Clientes c
        INNER JOIN dbo.NegocioClientes nc ON nc.ClienteId = c.Id
        WHERE c.Id = @Id
          AND nc.NegocioId = @NegocioId
          AND nc.Activo = 1;

        IF @@ROWCOUNT = 0
            RAISERROR('No se encontro el cliente para actualizar en el negocio.', 16, 1);

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar @NegocioId = @NegocioId, @Modulo = N'CLIENTES', @Accion = N'EDIT', @Entidad = N'Cliente', @EntidadId = @EntidadIdAudit, @Usuario = @Usuario, @DetalleJson = NULL;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
