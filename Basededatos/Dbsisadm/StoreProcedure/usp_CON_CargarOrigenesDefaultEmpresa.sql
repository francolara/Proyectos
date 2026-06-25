-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia origenes maestros internos hacia una empresa cuando aun no tiene origenes.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarOrigenesDefaultEmpresa
    @IdEmpresa INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa ya tiene origenes registrados.', 16, 1);
        END;

        INSERT INTO dbo.CON_Origen
        (
            IdEmpresa,
            CodigoOrigen,
            NombreOrigen,
            ModuloOrigen,
            PermiteRegistroManual,
            Estado,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            om.CodigoOrigen,
            om.NombreOrigen,
            om.ModuloOrigen,
            om.PermiteRegistroManual,
            om.Estado,
            @UsuarioRegistro
        FROM dbo.CON_OrigenMaestro AS om
        WHERE om.Estado = 1
        ORDER BY om.Orden, om.CodigoOrigen;

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
