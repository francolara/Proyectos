-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/04/2026
-- Description:   Seed de parametros globales para correos de notificacion en pestaña Portal Web (owner plataforma).
-- =============================================

IF NOT EXISTS (SELECT 1 FROM dbo.ParametrosGlobales WHERE NombreParametro = N'HOME_PORTAL_NOTIF_CORREO_1')
BEGIN
    INSERT INTO dbo.ParametrosGlobales
    (
        NombreParametro,
        Descripcion,
        ValorParametro,
        FechaCreacion,
        UsuarioCreacion
    )
    VALUES
    (
        N'HOME_PORTAL_NOTIF_CORREO_1',
        N'Correo principal para futuras notificaciones internas del portal web.',
        N'',
        SYSUTCDATETIME(),
        N'script'
    );
END;

IF NOT EXISTS (SELECT 1 FROM dbo.ParametrosGlobales WHERE NombreParametro = N'HOME_PORTAL_NOTIF_CORREO_2')
BEGIN
    INSERT INTO dbo.ParametrosGlobales
    (
        NombreParametro,
        Descripcion,
        ValorParametro,
        FechaCreacion,
        UsuarioCreacion
    )
    VALUES
    (
        N'HOME_PORTAL_NOTIF_CORREO_2',
        N'Correo secundario para futuras notificaciones internas del portal web.',
        N'',
        SYSUTCDATETIME(),
        N'script'
    );
END;
