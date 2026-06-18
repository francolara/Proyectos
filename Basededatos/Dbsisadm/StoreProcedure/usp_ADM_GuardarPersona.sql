-- =============================================
-- Author:        FRANCO LARA
-- Create date:   17/06/2026
-- Description:   Inserta o actualiza personas por empresa, sincroniza cliente/proveedor y asigna ubigeo 150101 por defecto en altas operativas.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_GuardarPersona
    @IdPersona INT = NULL,
    @IdEmpresa INT,
    @TipoPersona CHAR(1),
    @TipoDocumento VARCHAR(3),
    @NumeroDocumento VARCHAR(20),
    @ApellidoPaterno NVARCHAR(100) = NULL,
    @ApellidoMaterno NVARCHAR(100) = NULL,
    @Nombres NVARCHAR(150) = NULL,
    @RazonSocial NVARCHAR(200) = NULL,
    @CorreoElectronico NVARCHAR(200) = NULL,
    @Telefono NVARCHAR(50) = NULL,
    @Direccion NVARCHAR(250) = NULL,
    @CodigoUbigeo CHAR(6) = NULL,
    @EsCliente BIT,
    @EsProveedor BIT,
    @Estado BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdPersonaTrabajo INT
        DECLARE @CodigoCliente VARCHAR(20)
        DECLARE @CodigoProveedor VARCHAR(20)

        SET @TipoPersona = UPPER(LTRIM(RTRIM(@TipoPersona)));
        SET @TipoDocumento = UPPER(LTRIM(RTRIM(@TipoDocumento)));
        SET @NumeroDocumento = LTRIM(RTRIM(@NumeroDocumento));
        SET @ApellidoPaterno = NULLIF(LTRIM(RTRIM(@ApellidoPaterno)), N'');
        SET @ApellidoMaterno = NULLIF(LTRIM(RTRIM(@ApellidoMaterno)), N'');
        SET @Nombres = NULLIF(LTRIM(RTRIM(@Nombres)), N'');
        SET @RazonSocial = NULLIF(LTRIM(RTRIM(@RazonSocial)), N'');
        SET @CorreoElectronico = NULLIF(LTRIM(RTRIM(@CorreoElectronico)), N'');
        SET @Telefono = NULLIF(LTRIM(RTRIM(@Telefono)), N'');
        SET @Direccion = NULLIF(LTRIM(RTRIM(@Direccion)), N'');
        SET @CodigoUbigeo = NULLIF(LTRIM(RTRIM(@CodigoUbigeo)), '');
        SET @CodigoUbigeo = ISNULL(@CodigoUbigeo, '150101');

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa activa no existe.', 16, 1);
        END;

        IF @TipoPersona NOT IN ('N', 'J')
        BEGIN
            RAISERROR(N'El tipo de persona es invalido.', 16, 1);
        END;

        IF @NumeroDocumento IS NULL OR @NumeroDocumento = ''
        BEGIN
            RAISERROR(N'Debe ingresar el numero de documento.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.TiposDocumentoIdentidadSunat AS td
            WHERE td.CodigoSunat = @TipoDocumento
              AND td.Activo = 1
        )
        BEGIN
            RAISERROR(N'El tipo de documento no existe o esta inactivo.', 16, 1);
        END;

        IF @TipoPersona = 'N' AND @Nombres IS NULL
        BEGIN
            RAISERROR(N'Debe ingresar los nombres de la persona natural.', 16, 1);
        END;

        IF @TipoPersona = 'J' AND @RazonSocial IS NULL
        BEGIN
            RAISERROR(N'Debe ingresar la razon social de la persona juridica.', 16, 1);
        END;

        IF @CodigoUbigeo IS NOT NULL
           AND NOT EXISTS
           (
                SELECT 1
                FROM dbo.UbigeoDistritos AS u
                WHERE u.CodigoUbigeo = @CodigoUbigeo
                  AND u.Activo = 1
           )
        BEGIN
            RAISERROR(N'El ubigeo seleccionado no existe o esta inactivo.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.ADM_Persona AS p
            WHERE p.IdEmpresa = @IdEmpresa
              AND p.TipoDocumento = @TipoDocumento
              AND p.NumeroDocumento = @NumeroDocumento
              AND (@IdPersona IS NULL OR p.IdPersona <> @IdPersona)
        )
        BEGIN
            RAISERROR(N'Ya existe una persona con el mismo documento en la empresa activa.', 16, 1);
        END;

        BEGIN TRAN;

        IF @IdPersona IS NULL
        BEGIN
            INSERT INTO dbo.ADM_Persona
            (
                IdEmpresa,
                TipoPersona,
                TipoDocumento,
                NumeroDocumento,
                ApellidoPaterno,
                ApellidoMaterno,
                Nombres,
                RazonSocial,
                CorreoElectronico,
                Telefono,
                Direccion,
                CodigoUbigeo,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @TipoPersona,
                @TipoDocumento,
                @NumeroDocumento,
                @ApellidoPaterno,
                @ApellidoMaterno,
                @Nombres,
                @RazonSocial,
                @CorreoElectronico,
                @Telefono,
                @Direccion,
                @CodigoUbigeo,
                @Estado,
                @UsuarioRegistro
            );

            SET @IdPersonaTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.ADM_Persona AS p
                WHERE p.IdPersona = @IdPersona
                  AND p.IdEmpresa = @IdEmpresa
            )
            BEGIN
                RAISERROR(N'La persona a editar no pertenece a la empresa activa.', 16, 1);
            END;

            UPDATE dbo.ADM_Persona
            SET TipoPersona = @TipoPersona,
                TipoDocumento = @TipoDocumento,
                NumeroDocumento = @NumeroDocumento,
                ApellidoPaterno = @ApellidoPaterno,
                ApellidoMaterno = @ApellidoMaterno,
                Nombres = @Nombres,
                RazonSocial = @RazonSocial,
                CorreoElectronico = @CorreoElectronico,
                Telefono = @Telefono,
                Direccion = @Direccion,
                CodigoUbigeo = @CodigoUbigeo,
                Estado = @Estado,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdPersona = @IdPersona
              AND IdEmpresa = @IdEmpresa;

            SET @IdPersonaTrabajo = @IdPersona;
        END;

        IF @EsCliente = 1
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.ADM_Cliente AS c
                WHERE c.IdEmpresa = @IdEmpresa
                  AND c.IdPersona = @IdPersonaTrabajo
            )
            BEGIN
                UPDATE dbo.ADM_Cliente
                SET Estado = 1,
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdPersona = @IdPersonaTrabajo;
            END
            ELSE
            BEGIN
                SELECT
                    @CodigoCliente = CONCAT(
                        'CLI',
                        RIGHT(
                            '000000' + CONVERT(VARCHAR(6), ISNULL(MAX(TRY_CONVERT(INT, RIGHT(c.CodigoCliente, 6))), 0) + 1),
                            6
                        )
                    )
                FROM dbo.ADM_Cliente AS c
                WHERE c.IdEmpresa = @IdEmpresa;

                INSERT INTO dbo.ADM_Cliente
                (
                    IdEmpresa,
                    IdPersona,
                    CodigoCliente,
                    LimiteCredito,
                    DiasCredito,
                    Estado,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdPersonaTrabajo,
                    @CodigoCliente,
                    0,
                    0,
                    1,
                    @UsuarioRegistro
                );
            END;
        END
        ELSE
        BEGIN
            UPDATE dbo.ADM_Cliente
            SET Estado = 0,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdPersona = @IdPersonaTrabajo;
        END;

        IF @EsProveedor = 1
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.ADM_Proveedor AS p
                WHERE p.IdEmpresa = @IdEmpresa
                  AND p.IdPersona = @IdPersonaTrabajo
            )
            BEGIN
                UPDATE dbo.ADM_Proveedor
                SET Estado = 1,
                    UsuarioRegistro = @UsuarioRegistro
                WHERE IdEmpresa = @IdEmpresa
                  AND IdPersona = @IdPersonaTrabajo;
            END
            ELSE
            BEGIN
                SELECT
                    @CodigoProveedor = CONCAT(
                        'PRV',
                        RIGHT(
                            '000000' + CONVERT(VARCHAR(6), ISNULL(MAX(TRY_CONVERT(INT, RIGHT(p.CodigoProveedor, 6))), 0) + 1),
                            6
                        )
                    )
                FROM dbo.ADM_Proveedor AS p
                WHERE p.IdEmpresa = @IdEmpresa;

                INSERT INTO dbo.ADM_Proveedor
                (
                    IdEmpresa,
                    IdPersona,
                    CodigoProveedor,
                    Estado,
                    UsuarioRegistro
                )
                VALUES
                (
                    @IdEmpresa,
                    @IdPersonaTrabajo,
                    @CodigoProveedor,
                    1,
                    @UsuarioRegistro
                );
            END;
        END
        ELSE
        BEGIN
            UPDATE dbo.ADM_Proveedor
            SET Estado = 0,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdPersona = @IdPersonaTrabajo;
        END;

        COMMIT;

        EXEC dbo.usp_ADM_ObtenerPersona
            @IdEmpresa = @IdEmpresa,
            @IdPersona = @IdPersonaTrabajo;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK;
        END;

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
