-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Registra una nueva empresa dentro de una cuenta administradora existente y asigna el usuario administrador.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Agrega carga automatica de parametros base desde maestro interno al crear empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_RegistrarEmpresaCuentaAdministradora
    @IdCuentaAdministradora INT,
    @AspNetUserId NVARCHAR(450),
    @CodigoEmpresa VARCHAR(20),
    @RazonSocial NVARCHAR(200),
    @NombreComercial NVARCHAR(200) = NULL,
    @Ruc VARCHAR(11),
    @EsEmpresaPredeterminada BIT = 0,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdEmpresa INT

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_CuentaAdministradora AS ca
            WHERE ca.IdCuentaAdministradora = @IdCuentaAdministradora
              AND ca.Estado = 1
        )
        BEGIN
            RAISERROR(N'La cuenta administradora no existe o esta inactiva.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.CodigoEmpresa = @CodigoEmpresa
               OR e.Ruc = @Ruc
        )
        BEGIN
            RAISERROR(N'La empresa ya existe con el mismo codigo o RUC.', 16, 1);
        END;

        INSERT INTO dbo.SEG_Empresa
        (
            IdCuentaAdministradora,
            CodigoEmpresa,
            RazonSocial,
            NombreComercial,
            Ruc,
            Estado,
            UsuarioRegistro
        )
        VALUES
        (
            @IdCuentaAdministradora,
            @CodigoEmpresa,
            @RazonSocial,
            @NombreComercial,
            @Ruc,
            1,
            @UsuarioRegistro
        );

        SET @IdEmpresa = SCOPE_IDENTITY();

        EXEC dbo.usp_SEG_AsignarUsuarioEmpresa
            @AspNetUserId = @AspNetUserId,
            @IdEmpresa = @IdEmpresa,
            @EsEmpresaPredeterminada = @EsEmpresaPredeterminada,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_ADM_CargarParametrosDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        SELECT
            @IdEmpresa AS IdEmpresa;

    END TRY

    BEGIN CATCH

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
