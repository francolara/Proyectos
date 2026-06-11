-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/06/2026
-- Firma:         Lista el historial comercial de suscripcion por negocio para superadmin.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_NegociosSuscripcionMovimiento_ListarPorNegocio
    @NegocioId INT,
    @Top INT = 8
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @NegocioId IS NULL OR @NegocioId <= 0
            RAISERROR('Negocio invalido.', 16, 1);

        SET @Top = CASE WHEN ISNULL(@Top, 0) <= 0 THEN 8 ELSE @Top END;

        SELECT TOP (@Top)
            m.Id,
            m.TipoMovimiento,
            CAST(COALESCE(m.EstadoSuscripcionAnterior, 0) AS INT) AS EstadoSuscripcionAnterior,
            CAST(COALESCE(m.EstadoSuscripcionNuevo, 0) AS INT) AS EstadoSuscripcionNuevo,
            CAST(COALESCE(m.EsPruebaAnterior, 0) AS BIT) AS EsPruebaAnterior,
            CAST(COALESCE(m.EsPruebaNuevo, 0) AS BIT) AS EsPruebaNuevo,
            m.TipoCobroAnterior,
            m.TipoCobroNuevo,
            m.FechaInicioReferencia,
            m.FechaFinReferencia,
            CAST(COALESCE(m.DiasGracia, 0) AS INT) AS DiasGracia,
            CAST(COALESCE(m.DiasExtra, 0) AS INT) AS DiasExtra,
            m.Observacion,
            m.FechaCreacion,
            m.UsuarioCreacion
        FROM dbo.NegociosSuscripcionMovimiento m
        WHERE m.NegocioId = @NegocioId
        ORDER BY m.FechaCreacion DESC, m.Id DESC;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
