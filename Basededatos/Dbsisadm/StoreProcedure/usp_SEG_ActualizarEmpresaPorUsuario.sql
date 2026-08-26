-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   26/08/2026
-- Description:   Actualiza los datos de una empresa autorizada y sincroniza CodigoEmpresa con el RUC.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_SEG_ActualizarEmpresaPorUsuario
    @IdEmpresa INT,
    @AspNetUserId NVARCHAR(450),
    @RazonSocial NVARCHAR(200),
    @NombreComercial NVARCHAR(200) = NULL,
    @Ruc VARCHAR(11),
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @RazonSocial = LTRIM(RTRIM(@RazonSocial));
        SET @NombreComercial = NULLIF(LTRIM(RTRIM(@NombreComercial)), N'');
        SET @Ruc = LTRIM(RTRIM(@Ruc));

        IF @IdEmpresa <= 0
        BEGIN
            RAISERROR (N'La empresa indicada no es valida.', 16, 1);
        END;

        IF NULLIF(@RazonSocial, N'') IS NULL
        BEGIN
            RAISERROR (N'La razon social es obligatoria.', 16, 1);
        END;

        IF LEN(@Ruc) <> 11 OR @Ruc LIKE '%[^0-9]%'
        BEGIN
            RAISERROR (N'El RUC debe contener exactamente 11 digitos.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioEmpresa AS usuarioEmpresa
            INNER JOIN dbo.SEG_Empresa AS empresa
                ON empresa.IdEmpresa = usuarioEmpresa.IdEmpresa
            WHERE usuarioEmpresa.AspNetUserId = @AspNetUserId
              AND usuarioEmpresa.IdEmpresa = @IdEmpresa
              AND usuarioEmpresa.Estado = 1
              AND empresa.Estado = 1
        )
        BEGIN
            RAISERROR (N'No tiene acceso para editar la empresa indicada.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS empresa
            WHERE empresa.IdEmpresa <> @IdEmpresa
              AND (empresa.Ruc = @Ruc OR empresa.CodigoEmpresa = @Ruc)
        )
        BEGIN
            RAISERROR (N'El RUC ya se encuentra registrado en otra empresa.', 16, 1);
        END;

        UPDATE empresa
        SET empresa.CodigoEmpresa = @Ruc,
            empresa.RazonSocial = @RazonSocial,
            empresa.NombreComercial = COALESCE(@NombreComercial, @RazonSocial),
            empresa.Ruc = @Ruc,
            empresa.UsuarioRegistro = @UsuarioRegistro
        FROM dbo.SEG_Empresa AS empresa
        WHERE empresa.IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH;
END;
