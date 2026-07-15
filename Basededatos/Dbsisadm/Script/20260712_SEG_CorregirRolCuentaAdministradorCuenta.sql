-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   12/07/2026
-- Description:   Corrige el valor legacy ADMINISTRADOR hacia ADMINISTRADORCUENTA en defaults, datos existentes y alta inicial de cuenta administradora.
-- =============================================

UPDATE uca
SET uca.RolCuenta = N'ADMINISTRADORCUENTA',
    uca.UsuarioRegistro = COALESCE(uca.UsuarioRegistro, N'sistema')
FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
WHERE uca.RolCuenta = N'ADMINISTRADOR';

IF EXISTS
(
    SELECT 1
    FROM sys.default_constraints AS dc
    WHERE dc.name = N'DF_SEG_UsuarioCuentaAdministradora_RolCuenta'
      AND dc.parent_object_id = OBJECT_ID(N'dbo.SEG_UsuarioCuentaAdministradora')
)
BEGIN
    ALTER TABLE dbo.SEG_UsuarioCuentaAdministradora
        DROP CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_RolCuenta;
END;

ALTER TABLE dbo.SEG_UsuarioCuentaAdministradora
    ADD CONSTRAINT DF_SEG_UsuarioCuentaAdministradora_RolCuenta DEFAULT (N'ADMINISTRADORCUENTA') FOR RolCuenta;

EXEC(N'
CREATE OR ALTER PROCEDURE dbo.usp_SEG_RegistrarCuentaAdministradoraConEmpresa
    @AspNetUserId NVARCHAR(450),
    @NombreCompleto NVARCHAR(180),
    @Telefono NVARCHAR(30) = NULL,
    @CorreoReferencia NVARCHAR(256),
    @CodigoCuenta VARCHAR(20),
    @NombreCuenta NVARCHAR(200),
    @CodigoEmpresa VARCHAR(20),
    @RazonSocial NVARCHAR(200),
    @NombreComercial NVARCHAR(200) = NULL,
    @Ruc VARCHAR(11),
    @DiasPrueba INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdCuentaAdministradora INT
        DECLARE @IdEmpresa INT
        DECLARE @IdCuentaAdministradoraSuscripcion INT
        DECLARE @FechaInicioPrueba DATE = CAST(SYSDATETIME() AS DATE)
        DECLARE @FechaFinPrueba DATE = DATEADD(DAY, @DiasPrueba, CAST(SYSDATETIME() AS DATE))

        IF OBJECT_ID(N''dbo.usp_SEG_SeedSeguridadCuentaPermisosBase'', N''P'') IS NOT NULL
        BEGIN
            EXEC dbo.usp_SEG_SeedSeguridadCuentaPermisosBase
                @UsuarioRegistro = COALESCE(@UsuarioRegistro, @CorreoReferencia, N''sistema'');
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_CuentaAdministradora AS ca
            WHERE ca.CodigoCuenta = @CodigoCuenta
        )
        BEGIN
            RAISERROR(N''El codigo de cuenta ya existe.'', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.CodigoEmpresa = @CodigoEmpresa
               OR e.Ruc = @Ruc
        )
        BEGIN
            RAISERROR(N''La empresa ya existe con el mismo codigo o RUC.'', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioPerfil AS up
            WHERE up.AspNetUserId = @AspNetUserId
        )
        BEGIN
            UPDATE dbo.SEG_UsuarioPerfil
            SET NombreCompleto = @NombreCompleto,
                Telefono = @Telefono,
                CorreoReferencia = @CorreoReferencia,
                UsuarioRegistro = @UsuarioRegistro
            WHERE AspNetUserId = @AspNetUserId;
        END
        ELSE
        BEGIN
            INSERT INTO dbo.SEG_UsuarioPerfil
            (
                AspNetUserId,
                NombreCompleto,
                Telefono,
                CorreoReferencia,
                UsuarioRegistro
            )
            VALUES
            (
                @AspNetUserId,
                @NombreCompleto,
                @Telefono,
                @CorreoReferencia,
                @UsuarioRegistro
            );
        END;

        INSERT INTO dbo.SEG_CuentaAdministradora
        (
            CodigoCuenta,
            NombreCuenta,
            CorreoPrincipal,
            TelefonoPrincipal,
            Estado,
            UsuarioRegistro
        )
        VALUES
        (
            @CodigoCuenta,
            @NombreCuenta,
            @CorreoReferencia,
            @Telefono,
            1,
            @UsuarioRegistro
        );

        SET @IdCuentaAdministradora = SCOPE_IDENTITY();

        INSERT INTO dbo.SEG_UsuarioCuentaAdministradora
        (
            AspNetUserId,
            IdCuentaAdministradora,
            RolCuenta,
            EsCuentaPredeterminada,
            Estado,
            UsuarioRegistro
        )
        VALUES
        (
            @AspNetUserId,
            @IdCuentaAdministradora,
            N''ADMINISTRADORCUENTA'',
            1,
            1,
            @UsuarioRegistro
        );

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

        INSERT INTO dbo.SEG_UsuarioEmpresa
        (
            AspNetUserId,
            IdEmpresa,
            EsEmpresaPredeterminada,
            Estado,
            UsuarioRegistro
        )
        VALUES
        (
            @AspNetUserId,
            @IdEmpresa,
            1,
            1,
            @UsuarioRegistro
        );

        EXEC dbo.usp_ADM_CargarParametrosDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        EXEC dbo.usp_CON_CargarPlanCuentaDefaultEmpresa
            @IdEmpresa = @IdEmpresa,
            @UsuarioRegistro = @UsuarioRegistro;

        INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcion
        (
            IdCuentaAdministradora,
            TipoPlan,
            EstadoSuscripcion,
            EsPrueba,
            FechaInicioPrueba,
            FechaFinPrueba,
            Activo,
            EmpresasPermitidas,
            UsuariosPermitidos,
            Observacion,
            UsuarioRegistro
        )
        VALUES
        (
            @IdCuentaAdministradora,
            N''TRIAL'',
            N''TRIAL'',
            1,
            @FechaInicioPrueba,
            @FechaFinPrueba,
            1,
            3,
            3,
            N''Registro inicial automatico.'',
            @UsuarioRegistro
        );

        SET @IdCuentaAdministradoraSuscripcion = SCOPE_IDENTITY();

        INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcionMovimiento
        (
            IdCuentaAdministradora,
            IdCuentaAdministradoraSuscripcion,
            TipoMovimiento,
            TipoPlanNuevo,
            EstadoSuscripcionNuevo,
            EsPruebaNuevo,
            FechaInicioReferencia,
            FechaFinReferencia,
            EmpresasPermitidasNuevo,
            UsuariosPermitidosNuevo,
            Observacion,
            UsuarioRegistro
        )
        VALUES
        (
            @IdCuentaAdministradora,
            @IdCuentaAdministradoraSuscripcion,
            N''ALTA_INICIAL'',
            N''TRIAL'',
            N''TRIAL'',
            1,
            @FechaInicioPrueba,
            @FechaFinPrueba,
            3,
            3,
            N''Creacion inicial de cuenta administradora y empresa principal.'',
            @UsuarioRegistro
        );

        SELECT
            @IdCuentaAdministradora AS IdCuentaAdministradora,
            @IdEmpresa AS IdEmpresa,
            @FechaInicioPrueba AS FechaInicioPrueba,
            @FechaFinPrueba AS FechaFinPrueba;

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
');
