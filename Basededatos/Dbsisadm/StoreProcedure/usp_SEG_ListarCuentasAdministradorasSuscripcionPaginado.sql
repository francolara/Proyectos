-- =============================================
-- Author:        FRANCO LARA
-- Create date:   25/07/2026
-- Description:   Pagina suscriptores, calcula su estado efectivo de acceso y devuelve los indicadores del SuperAdmin.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarCuentasAdministradorasSuscripcionPaginado
    @TextoBusqueda NVARCHAR(200) = NULL,
    @EstadoFiltro NVARCHAR(20) = N'TODOS',
    @NumeroPagina INT = 1,
    @TamanoPagina INT = 10
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SET @TextoBusqueda = NULLIF(LTRIM(RTRIM(@TextoBusqueda)), N'');
        SET @EstadoFiltro = UPPER(ISNULL(NULLIF(LTRIM(RTRIM(@EstadoFiltro)), N''), N'TODOS'));
        SET @NumeroPagina = CASE WHEN @NumeroPagina < 1 THEN 1 ELSE @NumeroPagina END;
        SET @TamanoPagina = CASE WHEN @TamanoPagina < 1 OR @TamanoPagina > 100 THEN 10 ELSE @TamanoPagina END;

        SELECT
            ca.IdCuentaAdministradora,
            ca.CodigoCuenta,
            ca.NombreCuenta,
            ca.CorreoPrincipal,
            ca.TelefonoPrincipal,
            ca.Estado AS EstadoCuenta,
            principal.IdEmpresa AS IdEmpresaPrincipal,
            principal.CodigoEmpresa AS CodigoEmpresaPrincipal,
            principal.RazonSocial AS RazonSocialEmpresaPrincipal,
            principal.NombreComercial AS NombreComercialEmpresaPrincipal,
            principal.Ruc AS RucEmpresaPrincipal,
            empresas.CantidadEmpresas,
            cas.IdCuentaAdministradoraSuscripcion,
            cas.TipoPlan,
            cas.EstadoSuscripcion,
            CASE
                WHEN UPPER(ISNULL(cas.EstadoSuscripcion, N'')) = N'BAJA'
                    THEN N'BAJA'
                WHEN ISNULL(cas.Activo, 0) = 0
                     OR UPPER(ISNULL(cas.EstadoSuscripcion, N'')) = N'SUSPENDIDO'
                    THEN N'SUSPENDIDO'
                WHEN
                (
                    ISNULL(cas.EsPrueba, 0) = 1
                    OR UPPER(ISNULL(cas.TipoPlan, N'')) IN (N'TRIAL', N'GRATIS')
                    OR UPPER(ISNULL(cas.EstadoSuscripcion, N'')) = N'TRIAL'
                )
                AND
                (
                    cas.FechaFinPrueba IS NULL
                    OR CAST(GETDATE() AS DATE) > cas.FechaFinPrueba
                )
                    THEN N'VENCIDO'
                WHEN
                (
                    ISNULL(cas.EsPrueba, 0) = 1
                    OR UPPER(ISNULL(cas.TipoPlan, N'')) IN (N'TRIAL', N'GRATIS')
                    OR UPPER(ISNULL(cas.EstadoSuscripcion, N'')) = N'TRIAL'
                )
                    THEN N'TRIAL'
                WHEN cas.FechaFinPlan IS NULL
                    THEN N'SIN_VIGENCIA'
                WHEN CAST(GETDATE() AS DATE) <= cas.FechaFinPlan
                    THEN N'ACTIVO'
                WHEN CAST(GETDATE() AS DATE) <=
                    COALESCE
                    (
                        cas.FechaFinGracia,
                        DATEADD(DAY, CASE WHEN ISNULL(cas.DiasGracia, 0) < 0 THEN 0 ELSE ISNULL(cas.DiasGracia, 0) END, cas.FechaFinPlan)
                    )
                    THEN N'EN_GRACIA'
                ELSE N'VENCIDO'
            END AS EstadoAccesoEfectivo,
            cas.EsPrueba,
            cas.FechaInicioPrueba,
            cas.FechaFinPrueba,
            cas.FechaInicioPlan,
            cas.FechaFinPlan,
            cas.TipoCobro,
            cas.DiasGracia,
            cas.FechaFinGracia,
            cas.EmpresasPermitidas,
            cas.UsuariosPermitidos,
            cas.Activo,
            cas.Observacion,
            uca.AspNetUserId,
            up.NombreCompleto,
            up.Telefono,
            au.Email,
            ca.FechaRegistro AS FechaRegistroOrden
        INTO #CuentasBase
        FROM dbo.SEG_CuentaAdministradora AS ca
        OUTER APPLY
        (
            SELECT
                COUNT(1) AS CantidadEmpresas
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdCuentaAdministradora = ca.IdCuentaAdministradora
        ) AS empresas
        OUTER APPLY
        (
            SELECT TOP (1)
                e.IdEmpresa,
                e.CodigoEmpresa,
                e.RazonSocial,
                e.NombreComercial,
                e.Ruc
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdCuentaAdministradora = ca.IdCuentaAdministradora
            ORDER BY
                e.IdEmpresa ASC
        ) AS principal
        LEFT JOIN dbo.SEG_CuentaAdministradoraSuscripcion AS cas
            ON cas.IdCuentaAdministradora = ca.IdCuentaAdministradora
        LEFT JOIN dbo.SEG_UsuarioCuentaAdministradora AS uca
            ON uca.IdCuentaAdministradora = ca.IdCuentaAdministradora
           AND uca.EsCuentaPredeterminada = 1
           AND uca.Estado = 1
        LEFT JOIN dbo.SEG_UsuarioPerfil AS up
            ON up.AspNetUserId = uca.AspNetUserId
        LEFT JOIN dbo.AspNetUsers AS au
            ON au.Id = uca.AspNetUserId;

        DECLARE @TotalFiltrado INT;
        DECLARE @TotalPaginas INT;
        DECLARE @PaginaActual INT;

        SELECT
            @TotalFiltrado = COUNT(1)
        FROM #CuentasBase AS cb
        WHERE
            (
                @TextoBusqueda IS NULL
                OR cb.NombreCuenta LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.CodigoCuenta LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.CorreoPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.NombreCompleto LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.Email LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.RucEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.NombreComercialEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.RazonSocialEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
            )
            AND
            (
                @EstadoFiltro = N'TODOS'
                OR
                (
                    @EstadoFiltro = N'ACTIVAS'
                    AND cb.EstadoCuenta = 1
                    AND cb.EstadoAccesoEfectivo IN (N'ACTIVO', N'EN_GRACIA')
                )
                OR
                (
                    @EstadoFiltro = N'TRIAL'
                    AND cb.EstadoAccesoEfectivo = N'TRIAL'
                )
                OR
                (
                    @EstadoFiltro = N'SUSPENDIDAS'
                    AND
                    (
                        cb.EstadoCuenta = 0
                        OR cb.EstadoAccesoEfectivo IN (N'VENCIDO', N'SUSPENDIDO', N'BAJA', N'SIN_VIGENCIA')
                    )
                )
            );

        SET @TotalPaginas = CASE
            WHEN @TotalFiltrado = 0 THEN 1
            ELSE CEILING(@TotalFiltrado * 1.0 / @TamanoPagina)
        END;
        SET @PaginaActual = CASE
            WHEN @NumeroPagina > @TotalPaginas THEN @TotalPaginas
            ELSE @NumeroPagina
        END;

        SELECT
            @PaginaActual AS PaginaActual,
            @TamanoPagina AS TamanoPagina,
            @TotalFiltrado AS TotalFiltrado,
            @TotalPaginas AS TotalPaginas,
            (SELECT COUNT(1) FROM #CuentasBase) AS TotalCuentas,
            (
                SELECT COUNT(1)
                FROM #CuentasBase AS cb
                WHERE cb.EstadoCuenta = 1
                  AND cb.EstadoAccesoEfectivo IN (N'ACTIVO', N'EN_GRACIA')
            ) AS CuentasActivas,
            (
                SELECT COUNT(1)
                FROM #CuentasBase AS cb
                WHERE cb.EstadoCuenta = 1
                  AND cb.EstadoAccesoEfectivo = N'TRIAL'
            ) AS CuentasEnPrueba,
            (
                SELECT COUNT(1)
                FROM #CuentasBase AS cb
                WHERE cb.EstadoCuenta = 0
                   OR cb.EstadoAccesoEfectivo IN (N'VENCIDO', N'SUSPENDIDO', N'BAJA', N'SIN_VIGENCIA')
            ) AS CuentasSuspendidasOBaja,
            (
                SELECT COUNT(1)
                FROM dbo.SEG_CuentaAdministradoraSuscripcionPago AS p
            ) AS CobrosRegistrados,
            (
                SELECT COUNT(1)
                FROM dbo.SEG_CuentaAdministradoraSuscripcionPago AS p
                WHERE p.EstadoPago = N'PENDIENTE'
                  AND p.AplicarAlConfirmar = 1
                  AND ISNULL(p.AplicadoSuscripcion, 0) = 0
            ) AS CobrosPendientesAplicacion,
            CAST
            (
                ISNULL
                (
                    (
                        SELECT SUM(p.Monto)
                        FROM dbo.SEG_CuentaAdministradoraSuscripcionPago AS p
                        WHERE p.EstadoPago = N'PAGADO'
                          AND p.FechaPago >= DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1)
                    ),
                    0
                )
                AS DECIMAL(18, 2)
            ) AS MontoCobradoMes;

        SELECT
            cb.IdCuentaAdministradora,
            cb.CodigoCuenta,
            cb.NombreCuenta,
            cb.CorreoPrincipal,
            cb.TelefonoPrincipal,
            cb.EstadoCuenta,
            cb.IdEmpresaPrincipal,
            cb.CodigoEmpresaPrincipal,
            cb.RazonSocialEmpresaPrincipal,
            cb.NombreComercialEmpresaPrincipal,
            cb.RucEmpresaPrincipal,
            cb.CantidadEmpresas,
            cb.IdCuentaAdministradoraSuscripcion,
            cb.TipoPlan,
            cb.EstadoSuscripcion,
            cb.EsPrueba,
            cb.FechaInicioPrueba,
            cb.FechaFinPrueba,
            cb.FechaInicioPlan,
            cb.FechaFinPlan,
            cb.TipoCobro,
            cb.DiasGracia,
            cb.FechaFinGracia,
            cb.EmpresasPermitidas,
            cb.UsuariosPermitidos,
            cb.Activo,
            cb.Observacion,
            cb.AspNetUserId,
            cb.NombreCompleto,
            cb.Telefono,
            cb.Email
        FROM #CuentasBase AS cb
        WHERE
            (
                @TextoBusqueda IS NULL
                OR cb.NombreCuenta LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.CodigoCuenta LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.CorreoPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.NombreCompleto LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.Email LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.RucEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.NombreComercialEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
                OR cb.RazonSocialEmpresaPrincipal LIKE N'%' + @TextoBusqueda + N'%'
            )
            AND
            (
                @EstadoFiltro = N'TODOS'
                OR
                (
                    @EstadoFiltro = N'ACTIVAS'
                    AND cb.EstadoCuenta = 1
                    AND cb.EstadoAccesoEfectivo IN (N'ACTIVO', N'EN_GRACIA')
                )
                OR
                (
                    @EstadoFiltro = N'TRIAL'
                    AND cb.EstadoAccesoEfectivo = N'TRIAL'
                )
                OR
                (
                    @EstadoFiltro = N'SUSPENDIDAS'
                    AND
                    (
                        cb.EstadoCuenta = 0
                        OR cb.EstadoAccesoEfectivo IN (N'VENCIDO', N'SUSPENDIDO', N'BAJA', N'SIN_VIGENCIA')
                    )
                )
            )
        ORDER BY
            cb.FechaRegistroOrden DESC,
            cb.NombreCuenta ASC
        OFFSET (@PaginaActual - 1) * @TamanoPagina ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;

        DROP TABLE #CuentasBase;

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
