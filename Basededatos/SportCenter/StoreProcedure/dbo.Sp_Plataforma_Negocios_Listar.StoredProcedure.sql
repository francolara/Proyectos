-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Listado de negocios para panel superadmin con limites operativos y estado comercial.
-- Firma:         20/04/2026 | Agrega filtro comercial y paginacion backend con total de registros para panel superadmin.
-- Firma:         FRANCO LARA - 18/06/2026 | Expone TipoPlan para administrar capacidad comercial Basico/Full desde plataforma.
-- Firma:         FRANCO LARA - 21/07/2026 | Expone el plan comercial vigente y la fecha de registro para los reportes HTML del Super Admin.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocios_Listar
    @Buscar NVARCHAR(200) = NULL,
    @EstadoContrato NVARCHAR(30) = N'todos',
    @Pagina INT = 1,
    @TamanoPagina INT = 20,
    @TotalRegistros INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @EstadoContratoNorm NVARCHAR(30) = LOWER(LTRIM(RTRIM(COALESCE(@EstadoContrato, N'todos'))));
        DECLARE @Hoy DATE = CAST(GETDATE() AS DATE);

        SET @Pagina = CASE WHEN ISNULL(@Pagina, 0) < 1 THEN 1 ELSE @Pagina END;
        SET @TamanoPagina = CASE WHEN ISNULL(@TamanoPagina, 0) < 1 THEN 20 ELSE @TamanoPagina END;
        SET @Buscar = NULLIF(LTRIM(RTRIM(@Buscar)), N'');

        SELECT
            @TotalRegistros = COUNT(1)
        FROM dbo.Negocios n
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE (@Buscar IS NULL OR n.NombreComercial LIKE N'%' + @Buscar + N'%')
          AND (
                @EstadoContratoNorm = N'todos'
                OR (@EstadoContratoNorm = N'con-contrato' AND COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL)
                OR (@EstadoContratoNorm = N'sin-contrato' AND NOT (COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL))
                OR (@EstadoContratoNorm = N'prueba-por-vencer' AND COALESCE(ns.EsPrueba, 0) = 1 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPrueba >= @Hoy AND ns.FechaFinPrueba <= DATEADD(DAY, 7, @Hoy))
              );

        SELECT
            n.Id,
            n.NombreComercial,
            n.Activo,
            CAST(COALESCE(n.TipoPlan, N'Basico') AS NVARCHAR(20)) AS TipoPlan,
            CAST(COALESCE(n.SedesPermitidas, 2) AS INT) AS SedesPermitidas,
            CAST(COALESCE(n.EspaciosPermitidos, 6) AS INT) AS EspaciosPermitidos,
            CAST(COALESCE(n.UsuariosPermitidos, 3) AS INT) AS UsuariosPermitidos,
            CAST(COALESCE(ns.EstadoSuscripcion, 0) AS INT) AS EstadoSuscripcion,
            CAST(COALESCE(ns.EsPrueba, 0) AS BIT) AS EsPrueba,
            ns.FechaInicioPrueba,
            ns.FechaFinPrueba,
            ns.TipoCobro,
            ns.FechaInicioPlan,
            ns.FechaFinPlan,
            CAST(COALESCE(ns.DiasGracia, 5) AS INT) AS DiasGracia,
            ns.FechaFinGracia,
            CAST(COALESCE(NULLIF(ns.PlanComercial, N''), CASE WHEN COALESCE(ns.EsPrueba, 0) = 1 THEN N'PRUEBA' ELSE N'ESENCIAL' END) AS NVARCHAR(20)) AS PlanComercial,
            n.FechaRegistro
        FROM dbo.Negocios n
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE (@Buscar IS NULL OR n.NombreComercial LIKE N'%' + @Buscar + N'%')
          AND (
                @EstadoContratoNorm = N'todos'
                OR (@EstadoContratoNorm = N'con-contrato' AND COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL)
                OR (@EstadoContratoNorm = N'sin-contrato' AND NOT (COALESCE(ns.EsPrueba, 0) = 0 AND ns.TipoCobro IS NOT NULL AND ns.FechaFinPlan IS NOT NULL))
                OR (@EstadoContratoNorm = N'prueba-por-vencer' AND COALESCE(ns.EsPrueba, 0) = 1 AND ns.FechaFinPrueba IS NOT NULL AND ns.FechaFinPrueba >= @Hoy AND ns.FechaFinPrueba <= DATEADD(DAY, 7, @Hoy))
              )
        ORDER BY n.NombreComercial, n.Id
        OFFSET ((@Pagina - 1) * @TamanoPagina) ROWS
        FETCH NEXT @TamanoPagina ROWS ONLY;
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
    END CATCH
END
