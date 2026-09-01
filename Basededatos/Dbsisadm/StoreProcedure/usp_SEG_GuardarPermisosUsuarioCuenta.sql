-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   29/08/2026
-- Description:   Guarda en una sola transaccion los overrides de permisos generales de un usuario.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_GuardarPermisosUsuarioCuenta
    @IdUsuarioCuentaAdministradora INT,
    @PermisosXml XML,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
            WHERE uca.IdUsuarioCuentaAdministradora = @IdUsuarioCuentaAdministradora
              AND uca.Estado = 1
        )
        BEGIN
            RAISERROR (N'El usuario no pertenece a una cuenta administradora activa.', 16, 1);
            RETURN;
        END;

        DECLARE @Permisos TABLE
        (
            IdModuloSistema INT NOT NULL PRIMARY KEY,
            PuedeVer BIT NULL,
            PuedeCrear BIT NULL,
            PuedeEditar BIT NULL,
            PuedeEliminar BIT NULL
        );

        INSERT INTO @Permisos
        (
            IdModuloSistema,
            PuedeVer,
            PuedeCrear,
            PuedeEditar,
            PuedeEliminar
        )
        SELECT
            permiso.Nodo.value('(@IdModuloSistema)[1]', 'INT'),
            CASE permiso.Nodo.value('(@PuedeVer)[1]', 'CHAR(1)') WHEN '1' THEN 1 WHEN '0' THEN 0 ELSE NULL END,
            CASE permiso.Nodo.value('(@PuedeCrear)[1]', 'CHAR(1)') WHEN '1' THEN 1 WHEN '0' THEN 0 ELSE NULL END,
            CASE permiso.Nodo.value('(@PuedeEditar)[1]', 'CHAR(1)') WHEN '1' THEN 1 WHEN '0' THEN 0 ELSE NULL END,
            CASE permiso.Nodo.value('(@PuedeEliminar)[1]', 'CHAR(1)') WHEN '1' THEN 1 WHEN '0' THEN 0 ELSE NULL END
        FROM @PermisosXml.nodes('/Permisos/Permiso') AS permiso(Nodo);

        IF NOT EXISTS (SELECT 1 FROM @Permisos)
        BEGIN
            RAISERROR (N'No se recibieron permisos para guardar.', 16, 1);
            RETURN;
        END;

        IF EXISTS
        (
            SELECT 1
            FROM @Permisos AS p
            LEFT JOIN dbo.SEG_ModuloSistema AS ms
                ON ms.IdModuloSistema = p.IdModuloSistema
               AND ms.AlcanceModulo = 'CUENTA'
               AND ms.Estado = 1
            WHERE ms.IdModuloSistema IS NULL
        )
        BEGIN
            RAISERROR (N'Uno o mas modulos no pertenecen al alcance CUENTA o estan inactivos.', 16, 1);
            RETURN;
        END;

        BEGIN TRANSACTION;

        MERGE dbo.SEG_UsuarioCuentaPermiso AS destino
        USING
        (
            SELECT
                p.IdModuloSistema,
                p.PuedeVer,
                p.PuedeCrear,
                p.PuedeEditar,
                p.PuedeEliminar
            FROM @Permisos AS p
            WHERE p.PuedeVer IS NOT NULL
               OR p.PuedeCrear IS NOT NULL
               OR p.PuedeEditar IS NOT NULL
               OR p.PuedeEliminar IS NOT NULL
        ) AS origen
            ON destino.IdUsuarioCuentaAdministradora = @IdUsuarioCuentaAdministradora
           AND destino.IdModuloSistema = origen.IdModuloSistema
        WHEN MATCHED THEN
            UPDATE SET
                destino.PuedeVer = origen.PuedeVer,
                destino.PuedeCrear = origen.PuedeCrear,
                destino.PuedeEditar = origen.PuedeEditar,
                destino.PuedeEliminar = origen.PuedeEliminar,
                destino.Estado = 1,
                destino.UsuarioRegistro = @UsuarioRegistro
        WHEN NOT MATCHED BY TARGET THEN
            INSERT
            (
                IdUsuarioCuentaAdministradora,
                IdModuloSistema,
                PuedeVer,
                PuedeCrear,
                PuedeEditar,
                PuedeEliminar,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdUsuarioCuentaAdministradora,
                origen.IdModuloSistema,
                origen.PuedeVer,
                origen.PuedeCrear,
                origen.PuedeEditar,
                origen.PuedeEliminar,
                1,
                @UsuarioRegistro
            );

        DELETE permisoUsuario
        FROM dbo.SEG_UsuarioCuentaPermiso AS permisoUsuario
        INNER JOIN @Permisos AS permisoEntrada
            ON permisoEntrada.IdModuloSistema = permisoUsuario.IdModuloSistema
        WHERE permisoUsuario.IdUsuarioCuentaAdministradora = @IdUsuarioCuentaAdministradora
          AND permisoEntrada.PuedeVer IS NULL
          AND permisoEntrada.PuedeCrear IS NULL
          AND permisoEntrada.PuedeEditar IS NULL
          AND permisoEntrada.PuedeEliminar IS NULL;

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
