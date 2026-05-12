-- =============================================
-- Author:        FRANCO LARA
-- Create date:   12/05/2026
-- Firma:         Limpieza integral de datos de negocio y usuarios (excepto superadmin),
--                conservando ModulosSistema, incluyendo desafios, usuarios publicos,
--                referencias externas y tablas operativas adicionales con reinicio de IDENTITY.
-- =============================================

SET NOCOUNT ON;
SET XACT_ABORT ON;

BEGIN TRY
    BEGIN TRAN;

    DECLARE @SuperAdmins TABLE (UserId NVARCHAR(450) NOT NULL PRIMARY KEY);

    INSERT INTO @SuperAdmins (UserId)
    SELECT DISTINCT UR.UserId
    FROM dbo.AspNetUserRoles AS UR
    INNER JOIN dbo.AspNetRoles AS R
        ON R.Id = UR.RoleId
    WHERE R.NormalizedName = N'OWNERPLATAFORMA'
       OR R.Name = N'OwnerPlataforma';

    /* 1) Limpiar tablas hijas de negocio (orden por dependencias FK) */
    DELETE FROM dbo.DesafioMensaje;
    DELETE FROM dbo.Desafio;
    DELETE FROM dbo.ReservasUsuariosPublicos;
    DELETE FROM dbo.CuponesUso;
    DELETE FROM dbo.SolicitudesReservaPublica;
    DELETE FROM dbo.ComprobantesDetalle;
    DELETE FROM dbo.Pagos;
    DELETE FROM dbo.ComprobantesElectronicos;
    DELETE FROM dbo.Reservas;

    DELETE FROM dbo.BloqueosHorario;
    DELETE FROM dbo.PromocionesHorario;
    DELETE FROM dbo.Tarifas;

    DELETE FROM dbo.UsuariosPublicosPerfil;
    DELETE FROM dbo.HomeEspaciosReferencialesExternos;
    DELETE FROM dbo.Cupones;
    DELETE FROM dbo.PopupPromocion;

    DELETE FROM dbo.SedeConfiguracionNotificacion;
    DELETE FROM dbo.SedeFechasInhabilitadas;
    DELETE FROM dbo.SedeHorarioAtencion;
    DELETE FROM dbo.SedeServicios;
    DELETE FROM dbo.SedesSeriesDocumentoComprobante;

    DELETE FROM dbo.UsuariosNegocioPermiso;
    DELETE FROM dbo.UsuariosNegocio;
    --DELETE FROM dbo.RolesNegocioPermiso;

    DELETE FROM dbo.NegocioNotificaciones;
    DELETE FROM dbo.FormasPago;
    DELETE FROM dbo.NegociosFacturacionProveedorCredencial;
    DELETE FROM dbo.NegociosFacturacionProveedorConfig;
    DELETE FROM dbo.NegociosSeriesDocumentoComprobante;
    DELETE FROM dbo.NegociosSuscripcion;
    DELETE FROM dbo.NegociosTiposDocumentoComprobante;

    /* Romper FK circular entre Negocios y Monedas antes de eliminar */
    UPDATE dbo.Negocios
    SET MonedaId = NULL
    WHERE MonedaId IS NOT NULL;

    DELETE FROM dbo.BitacoraAuditoria;
    DELETE FROM dbo.SolicitudesAltaClub;
    DELETE FROM dbo.EspaciosDeportivos;
    DELETE FROM dbo.Sedes;
    DELETE FROM dbo.Clientes;

    /* Catalogos asociados al negocio (dependen de Negocios) */
    DELETE FROM dbo.TiposDeporte;
    DELETE FROM dbo.TiposSuelo;

    DELETE FROM dbo.Monedas;
    DELETE FROM dbo.Negocios;

    /* 2) Limpiar identidad/seguridad, conservando superadmin */
    IF EXISTS (SELECT 1 FROM @SuperAdmins)
    BEGIN
        DELETE UR
        FROM dbo.AspNetUserRoles AS UR
        WHERE NOT EXISTS (
            SELECT 1
            FROM @SuperAdmins AS SA
            WHERE SA.UserId = UR.UserId
        );

        DELETE UC
        FROM dbo.AspNetUserClaims AS UC
        WHERE NOT EXISTS (
            SELECT 1
            FROM @SuperAdmins AS SA
            WHERE SA.UserId = UC.UserId
        );

        DELETE UL
        FROM dbo.AspNetUserLogins AS UL
        WHERE NOT EXISTS (
            SELECT 1
            FROM @SuperAdmins AS SA
            WHERE SA.UserId = UL.UserId
        );

        DELETE UT
        FROM dbo.AspNetUserTokens AS UT
        WHERE NOT EXISTS (
            SELECT 1
            FROM @SuperAdmins AS SA
            WHERE SA.UserId = UT.UserId
        );

        DELETE U
        FROM dbo.AspNetUsers AS U
        WHERE NOT EXISTS (
            SELECT 1
            FROM @SuperAdmins AS SA
            WHERE SA.UserId = U.Id
        );
    END
    ELSE
    BEGIN
        PRINT 'No se encontro rol de superadmin/OwnerPlataforma; se omite limpieza de AspNetUsers.';
    END

    /* 3) Reiniciar IDENTITY de tablas limpiadas (solo si tienen IDENTITY) */
    DECLARE @TablasReseed TABLE (TableName SYSNAME NOT NULL PRIMARY KEY);

    INSERT INTO @TablasReseed (TableName)
    VALUES
        (N'DesafioMensaje'),
        (N'Desafio'),
        (N'ReservasUsuariosPublicos'),
        (N'CuponesUso'),
        (N'SolicitudesReservaPublica'),
        (N'ComprobantesDetalle'),
        (N'Pagos'),
        (N'ComprobantesElectronicos'),
        (N'Reservas'),
        (N'BloqueosHorario'),
        (N'PromocionesHorario'),
        (N'Tarifas'),
        (N'UsuariosPublicosPerfil'),
        (N'HomeEspaciosReferencialesExternos'),
        (N'Cupones'),
        (N'PopupPromocion'),
        (N'SedeConfiguracionNotificacion'),
        (N'SedeFechasInhabilitadas'),
        (N'SedeHorarioAtencion'),
        (N'SedeServicios'),
        (N'SedesSeriesDocumentoComprobante'),
        (N'UsuariosNegocioPermiso'),
        (N'UsuariosNegocio'),
        --(N'RolesNegocioPermiso'),
        (N'NegocioNotificaciones'),
        (N'FormasPago'),
        (N'Monedas'),
        (N'NegociosFacturacionProveedorCredencial'),
        (N'NegociosFacturacionProveedorConfig'),
        (N'NegociosSeriesDocumentoComprobante'),
        (N'NegociosSuscripcion'),
        (N'NegociosTiposDocumentoComprobante'),
        (N'BitacoraAuditoria'),
        (N'SolicitudesAltaClub'),
        (N'EspaciosDeportivos'),
        (N'Sedes'),
        (N'Clientes'),
        (N'Negocios'),
        (N'TiposDeporte'),
        (N'TiposSuelo');

    DECLARE @TableName SYSNAME;
    DECLARE @Sql NVARCHAR(500);

    DECLARE CurReseed CURSOR LOCAL FAST_FORWARD FOR
        SELECT TR.TableName
        FROM @TablasReseed AS TR
        INNER JOIN sys.identity_columns AS IC
            ON IC.object_id = OBJECT_ID(N'dbo.' + TR.TableName)
        GROUP BY TR.TableName;

    OPEN CurReseed;

    FETCH NEXT FROM CurReseed INTO @TableName;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        SET @Sql = N'DBCC CHECKIDENT (''dbo.' + @TableName + N''', RESEED, 0) WITH NO_INFOMSGS;';
        EXEC sp_executesql @Sql;

        FETCH NEXT FROM CurReseed INTO @TableName;
    END

    CLOSE CurReseed;
    DEALLOCATE CurReseed;

    COMMIT TRAN;
END TRY
BEGIN CATCH
    IF @@TRANCOUNT > 0
        ROLLBACK TRAN;

    DECLARE @ErrorMessage NVARCHAR(4000),
            @ErrorSeverity INT,
            @ErrorState INT;

    SELECT
        @ErrorMessage = ERROR_MESSAGE(),
        @ErrorSeverity = ERROR_SEVERITY(),
        @ErrorState = ERROR_STATE();

    RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
END CATCH;
