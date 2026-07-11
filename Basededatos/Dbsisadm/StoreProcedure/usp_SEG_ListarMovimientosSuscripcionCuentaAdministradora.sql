-- =============================================
-- Author:        FRANCO LARA
-- Create date:   10/07/2026
-- Description:   Lista el historial comercial de una cuenta administradora.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_SEG_ListarMovimientosSuscripcionCuentaAdministradora
    @IdCuentaAdministradora INT,
    @Top INT = 20
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SELECT TOP (CASE WHEN @Top <= 0 THEN 20 ELSE @Top END)
            m.IdCuentaAdministradoraSuscripcionMovimiento,
            m.TipoMovimiento,
            m.TipoPlanAnterior,
            m.TipoPlanNuevo,
            m.EstadoSuscripcionAnterior,
            m.EstadoSuscripcionNuevo,
            m.EsPruebaAnterior,
            m.EsPruebaNuevo,
            m.TipoCobroAnterior,
            m.TipoCobroNuevo,
            m.FechaInicioReferencia,
            m.FechaFinReferencia,
            m.DiasGracia,
            m.DiasExtra,
            m.EmpresasPermitidasAnterior,
            m.EmpresasPermitidasNuevo,
            m.UsuariosPermitidosAnterior,
            m.UsuariosPermitidosNuevo,
            m.Observacion,
            m.FechaRegistro,
            m.UsuarioRegistro
        FROM dbo.SEG_CuentaAdministradoraSuscripcionMovimiento AS m
        WHERE m.IdCuentaAdministradora = @IdCuentaAdministradora
        ORDER BY m.FechaRegistro DESC, m.IdCuentaAdministradoraSuscripcionMovimiento DESC;

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
