-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Migra la suscripcion de empresa hacia cuenta administradora y enlaza SEG_Empresa con su cuenta titular.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   12/07/2026
-- Description:   Corrige la migracion legacy para que el rol inicial insertado en SEG_UsuarioCuentaAdministradora sea ADMINISTRADORCUENTA.
-- =============================================

IF COL_LENGTH(N'dbo.SEG_Empresa', N'IdCuentaAdministradora') IS NULL
BEGIN
    ALTER TABLE dbo.SEG_Empresa
        ADD IdCuentaAdministradora INT NULL;
END;

IF EXISTS
(
    SELECT 1
    FROM sys.columns
    WHERE object_id = OBJECT_ID(N'dbo.SEG_Empresa')
      AND name = N'CodigoEmpresa'
      AND max_length < 20
)
BEGIN
    ALTER TABLE dbo.SEG_Empresa
        ALTER COLUMN CodigoEmpresa VARCHAR(20) NOT NULL;
END;

IF OBJECT_ID(N'dbo.SEG_CuentaAdministradora', N'U') IS NOT NULL
BEGIN
    INSERT INTO dbo.SEG_CuentaAdministradora
    (
        CodigoCuenta,
        NombreCuenta,
        CorreoPrincipal,
        TelefonoPrincipal,
        Estado,
        UsuarioRegistro
    )
    SELECT
        CONCAT(N'CTA-', e.CodigoEmpresa),
        COALESCE(e.NombreComercial, e.RazonSocial),
        COALESCE(up.CorreoReferencia, au.Email, CONCAT(N'cuenta', e.IdEmpresa, N'@pendiente.local')),
        up.Telefono,
        e.Estado,
        e.UsuarioRegistro
    FROM dbo.SEG_Empresa AS e
    OUTER APPLY
    (
        SELECT TOP (1)
            ue.AspNetUserId
        FROM dbo.SEG_UsuarioEmpresa AS ue
        WHERE ue.IdEmpresa = e.IdEmpresa
        ORDER BY
            ue.EsEmpresaPredeterminada DESC,
            ue.IdUsuarioEmpresa ASC
    ) AS propietario
    LEFT JOIN dbo.SEG_UsuarioPerfil AS up
        ON up.AspNetUserId = propietario.AspNetUserId
    LEFT JOIN dbo.AspNetUsers AS au
        ON au.Id = propietario.AspNetUserId
    WHERE e.IdCuentaAdministradora IS NULL
      AND NOT EXISTS
      (
          SELECT 1
          FROM dbo.SEG_CuentaAdministradora AS ca
          WHERE ca.CodigoCuenta = CONCAT(N'CTA-', e.CodigoEmpresa)
      );

    UPDATE e
    SET e.IdCuentaAdministradora = ca.IdCuentaAdministradora
    FROM dbo.SEG_Empresa AS e
    INNER JOIN dbo.SEG_CuentaAdministradora AS ca
        ON ca.CodigoCuenta = CONCAT(N'CTA-', e.CodigoEmpresa)
    WHERE e.IdCuentaAdministradora IS NULL;

    IF OBJECT_ID(N'dbo.SEG_UsuarioCuentaAdministradora', N'U') IS NOT NULL
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
        SELECT
            ue.AspNetUserId,
            e.IdCuentaAdministradora,
            N'ADMINISTRADORCUENTA',
            CASE WHEN ue.EsEmpresaPredeterminada = 1 THEN 1 ELSE 0 END,
            ue.Estado,
            ue.UsuarioRegistro
        FROM dbo.SEG_UsuarioEmpresa AS ue
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = ue.IdEmpresa
        WHERE e.IdCuentaAdministradora IS NOT NULL
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.SEG_UsuarioCuentaAdministradora AS uca
              WHERE uca.AspNetUserId = ue.AspNetUserId
                AND uca.IdCuentaAdministradora = e.IdCuentaAdministradora
          );
    END;

    IF OBJECT_ID(N'dbo.SEG_EmpresaSuscripcion', N'U') IS NOT NULL
       AND OBJECT_ID(N'dbo.SEG_CuentaAdministradoraSuscripcion', N'U') IS NOT NULL
    BEGIN
        INSERT INTO dbo.SEG_CuentaAdministradoraSuscripcion
        (
            IdCuentaAdministradora,
            TipoPlan,
            EstadoSuscripcion,
            EsPrueba,
            FechaInicioPrueba,
            FechaFinPrueba,
            FechaInicioPlan,
            FechaFinPlan,
            Activo,
            Observacion,
            FechaRegistro,
            UsuarioRegistro
        )
        SELECT
            e.IdCuentaAdministradora,
            es.TipoPlan,
            es.EstadoSuscripcion,
            es.EsPrueba,
            es.FechaInicioPrueba,
            es.FechaFinPrueba,
            es.FechaInicioPlan,
            es.FechaFinPlan,
            es.Activo,
            es.Observacion,
            es.FechaRegistro,
            es.UsuarioRegistro
        FROM dbo.SEG_EmpresaSuscripcion AS es
        INNER JOIN dbo.SEG_Empresa AS e
            ON e.IdEmpresa = es.IdEmpresa
        WHERE e.IdCuentaAdministradora IS NOT NULL
          AND NOT EXISTS
          (
              SELECT 1
              FROM dbo.SEG_CuentaAdministradoraSuscripcion AS cas
              WHERE cas.IdCuentaAdministradora = e.IdCuentaAdministradora
          );
    END;

    IF NOT EXISTS
    (
        SELECT 1
        FROM sys.foreign_keys
        WHERE name = N'FK_SEG_Empresa_SEG_CuentaAdministradora'
    )
       AND NOT EXISTS
       (
           SELECT 1
           FROM dbo.SEG_Empresa
           WHERE IdCuentaAdministradora IS NULL
       )
    BEGIN
        ALTER TABLE dbo.SEG_Empresa
            ALTER COLUMN IdCuentaAdministradora INT NOT NULL;

        ALTER TABLE dbo.SEG_Empresa
            ADD CONSTRAINT FK_SEG_Empresa_SEG_CuentaAdministradora
            FOREIGN KEY (IdCuentaAdministradora) REFERENCES dbo.SEG_CuentaAdministradora (IdCuentaAdministradora);
    END;
END;
