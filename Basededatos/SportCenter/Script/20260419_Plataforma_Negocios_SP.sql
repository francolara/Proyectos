-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Despliegue de SP para listar negocios y actualizar limites desde panel superadmin.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocios_Listar
    @Buscar NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            n.Id,
            n.NombreComercial,
            n.Activo,
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
            ns.FechaFinGracia
        FROM dbo.Negocios n
        LEFT JOIN dbo.NegociosSuscripcion ns ON ns.NegocioId = n.Id
        WHERE (@Buscar IS NULL OR n.NombreComercial LIKE N'%' + LTRIM(RTRIM(@Buscar)) + N'%')
        ORDER BY n.NombreComercial, n.Id;
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

CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocios_ActualizarLimites
    @NegocioId INT,
    @SedesPermitidas INT,
    @EspaciosPermitidos INT,
    @UsuariosPermitidos INT,
    @Usuario NVARCHAR(200) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId)
            RAISERROR('Negocio no encontrado.', 16, 1);

        UPDATE dbo.Negocios
        SET SedesPermitidas = CASE WHEN @SedesPermitidas < 1 THEN 1 ELSE @SedesPermitidas END,
            EspaciosPermitidos = CASE WHEN @EspaciosPermitidos < 1 THEN 1 ELSE @EspaciosPermitidos END,
            UsuariosPermitidos = CASE WHEN @UsuariosPermitidos < 1 THEN 1 ELSE @UsuariosPermitidos END
        WHERE Id = @NegocioId;
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
