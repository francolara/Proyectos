-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   10/07/2026
-- Description:   Asigna o reactiva la relacion entre usuario y cuenta administradora validando el rol de cuenta.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   25/07/2026
-- Description:   Valida dentro de la transaccion el limite efectivo de usuarios configurado en la suscripcion antes de crear o reactivar accesos.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_AsignarUsuarioCuentaAdministradora
    @AspNetUserId NVARCHAR(450),
    @IdCuentaAdministradora INT,
    @RolCuenta VARCHAR(30),
    @EsCuentaPredeterminada BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    BEGIN TRY

        DECLARE @EstadoAccesoExistente BIT;
        DECLARE @UsuariosPermitidos INT;
        DECLARE @UsuariosActivos INT;
        DECLARE @RequiereCupo BIT;

        SET @RolCuenta = UPPER(LTRIM(RTRIM(@RolCuenta)));

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_RolCuenta AS rc
            WHERE rc.CodigoRolCuenta = @RolCuenta
              AND rc.Estado = 1
        )
        BEGIN
            RAISERROR (N'El rol de cuenta no existe o no esta activo.', 16, 1);
            RETURN;
        END;

        BEGIN TRANSACTION;

        SELECT
            @EstadoAccesoExistente = uca.Estado
        FROM dbo.SEG_UsuarioCuentaAdministradora AS uca WITH (UPDLOCK, HOLDLOCK)
        WHERE uca.AspNetUserId = @AspNetUserId
          AND uca.IdCuentaAdministradora = @IdCuentaAdministradora;

        SET @RequiereCupo = CASE
            WHEN @EstadoAccesoExistente = 1 THEN 0
            ELSE 1
        END;

        IF @RequiereCupo = 1
        BEGIN
            SELECT
                @UsuariosPermitidos = cas.UsuariosPermitidos
            FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas WITH (UPDLOCK, HOLDLOCK)
            WHERE cas.IdCuentaAdministradora = @IdCuentaAdministradora;

            IF @@ROWCOUNT = 0
            BEGIN
                RAISERROR(N'La cuenta administradora no tiene una suscripcion configurada.', 16, 1);
            END;

            IF @UsuariosPermitidos IS NOT NULL AND @UsuariosPermitidos <= 0
            BEGIN
                RAISERROR(N'El limite de usuarios de la suscripcion no tiene una configuracion valida.', 16, 1);
            END;

            SELECT
                @UsuariosActivos = COUNT(1)
            FROM dbo.SEG_UsuarioCuentaAdministradora AS uca WITH (UPDLOCK, HOLDLOCK)
            WHERE uca.IdCuentaAdministradora = @IdCuentaAdministradora
              AND uca.Estado = 1;

            IF @UsuariosPermitidos IS NOT NULL
               AND @UsuariosActivos >= @UsuariosPermitidos
            BEGIN
                RAISERROR(
                    N'La cuenta alcanzo el limite de %d usuario(s) permitido por su suscripcion.',
                    16,
                    1,
                    @UsuariosPermitidos);
            END;
        END;

        IF @EsCuentaPredeterminada = 1
        BEGIN
            UPDATE dbo.SEG_UsuarioCuentaAdministradora
            SET EsCuentaPredeterminada = 0
            WHERE AspNetUserId = @AspNetUserId;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
            WHERE uca.AspNetUserId = @AspNetUserId
              AND uca.IdCuentaAdministradora = @IdCuentaAdministradora
        )
        BEGIN
            UPDATE dbo.SEG_UsuarioCuentaAdministradora
            SET RolCuenta = @RolCuenta,
                EsCuentaPredeterminada = @EsCuentaPredeterminada,
                Estado = 1,
                UsuarioRegistro = @UsuarioRegistro
            WHERE AspNetUserId = @AspNetUserId
              AND IdCuentaAdministradora = @IdCuentaAdministradora;
        END
        ELSE
        BEGIN
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
                @RolCuenta,
                @EsCuentaPredeterminada,
                1,
                @UsuarioRegistro
            );
        END;

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF XACT_STATE() <> 0
        BEGIN
            ROLLBACK TRANSACTION;
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
