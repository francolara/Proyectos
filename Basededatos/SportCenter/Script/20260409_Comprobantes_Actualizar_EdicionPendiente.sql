/*
Firma: Codex - 09/04/2026
Descripcion: Actualiza Sp_Comprobantes_Actualizar para permitir solo edicion de datos del cliente cuando el comprobante esta pendiente.
*/
USE [DbSportCenter]
GO

CREATE OR ALTER PROCEDURE [dbo].[Sp_Comprobantes_Actualizar]
    @Id INT,
    @NegocioId INT,
    @ClienteCorreo NVARCHAR(200) = NULL,
    @ClienteTipoDocumento NVARCHAR(20) = NULL,
    @ClienteNumeroDocumento NVARCHAR(20) = NULL,
    @ClienteDireccionFiscal NVARCHAR(250) = NULL,
    @ClienteCodigoUbigeo CHAR(6) = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @ClienteId INT;
        DECLARE @EstadoActual INT;
        DECLARE @TipoComprobante INT;
        DECLARE @Total DECIMAL(10,2);
        DECLARE @TipoDocumentoFinal NVARCHAR(20);
        DECLARE @NumeroDocumentoFinal NVARCHAR(20);
        DECLARE @DireccionFinal NVARCHAR(250);
        DECLARE @UbigeoFinal CHAR(6);
        DECLARE @CodigoDocComprobante NVARCHAR(4);

        SELECT
            @ClienteId = c.ClienteId,
            @EstadoActual = c.Estado,
            @TipoComprobante = c.TipoComprobante,
            @Total = c.Total
        FROM dbo.ComprobantesElectronicos c
        WHERE c.Id = @Id
          AND c.NegocioId = @NegocioId;

        IF @ClienteId IS NULL
            RAISERROR('No se encontro el comprobante para actualizar en el negocio.', 16, 1);

        IF @EstadoActual <> 1
            RAISERROR('Solo se permite editar datos del cliente cuando el comprobante esta pendiente.', 16, 1);

        SET @CodigoDocComprobante =
            CASE
                WHEN @TipoComprobante = 2 THEN N'01'
                WHEN @TipoComprobante = 1 THEN N'03'
                WHEN @TipoComprobante = 3 THEN N'RI'
                ELSE N'03'
            END;

        SET @ClienteCorreo = NULLIF(LTRIM(RTRIM(@ClienteCorreo)), N'');
        SET @ClienteTipoDocumento = NULLIF(LTRIM(RTRIM(@ClienteTipoDocumento)), N'');
        SET @ClienteNumeroDocumento = NULLIF(LTRIM(RTRIM(@ClienteNumeroDocumento)), N'');
        SET @ClienteDireccionFiscal = NULLIF(LTRIM(RTRIM(@ClienteDireccionFiscal)), N'');
        SET @ClienteCodigoUbigeo = NULLIF(LTRIM(RTRIM(@ClienteCodigoUbigeo)), '');

        SELECT @TipoDocumentoFinal = ISNULL(@ClienteTipoDocumento, c.TipoDocumento)
        FROM dbo.Clientes c
        WHERE c.Id = @ClienteId
          AND c.NegocioId = @NegocioId;

        SET @NumeroDocumentoFinal = @ClienteNumeroDocumento;
        SET @DireccionFinal = @ClienteDireccionFiscal;
        SET @UbigeoFinal = @ClienteCodigoUbigeo;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.TiposDocumentoIdentidadSunat t
            WHERE t.CodigoSunat = @TipoDocumentoFinal
              AND t.Activo = 1
        )
            RAISERROR('El tipo de documento del cliente no es valido.', 16, 1);

        IF @CodigoDocComprobante = N'01' AND @TipoDocumentoFinal <> N'6'
            RAISERROR('Para Factura el cliente debe tener documento RUC (6).', 16, 1);

        IF @CodigoDocComprobante = N'03'
        BEGIN
            IF @Total > 700
                RAISERROR('La boleta no puede exceder S/ 700.00.', 16, 1);

            IF @TipoDocumentoFinal NOT IN (N'0', N'1')
                RAISERROR('Para Boleta el cliente debe tener tipo de documento 0 o 1.', 16, 1);
        END

        IF @TipoDocumentoFinal = N'0'
            SET @NumeroDocumentoFinal = NULL;
        ELSE IF @NumeroDocumentoFinal IS NULL
            RAISERROR('Ingresa el numero de documento del cliente.', 16, 1);

        IF @DireccionFinal IS NULL
            SET @UbigeoFinal = NULL;

        IF @DireccionFinal IS NOT NULL AND @UbigeoFinal IS NULL
            RAISERROR('Cuando se ingresa direccion, el ubigeo es obligatorio.', 16, 1);

        IF @UbigeoFinal IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.UbigeoDistritos WHERE CodigoUbigeo = @UbigeoFinal AND Activo = 1)
            RAISERROR('El codigo de ubigeo del cliente no existe.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.Clientes
        SET
            Correo = @ClienteCorreo,
            TipoDocumento = @TipoDocumentoFinal,
            NumeroDocumento = ISNULL(@NumeroDocumentoFinal, N''),
            DireccionFiscal = @DireccionFinal,
            CodigoUbigeo = @UbigeoFinal,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @ClienteId
          AND NegocioId = @NegocioId;

        UPDATE dbo.ComprobantesElectronicos
        SET FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE Id = @Id
          AND NegocioId = @NegocioId;

        DECLARE @EntidadIdAudit NVARCHAR(80);
        SET @EntidadIdAudit = CONVERT(NVARCHAR(80), @Id);
        EXEC dbo.Sp_Auditoria_Registrar
            @NegocioId = @NegocioId,
            @Modulo = N'COMPROBANTES',
            @Accion = N'EDIT',
            @Entidad = N'ComprobanteElectronico',
            @EntidadId = @EntidadIdAudit,
            @Usuario = @Usuario,
            @DetalleJson = NULL;

        COMMIT TRANSACTION;
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
