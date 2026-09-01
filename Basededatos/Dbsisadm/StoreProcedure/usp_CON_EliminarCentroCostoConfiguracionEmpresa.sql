-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/08/2026
-- Description:   Elimina un centro de costo por empresa solo cuando su codigo no fue usado en detalles contables o bancarios.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.usp_CON_EliminarCentroCostoConfiguracionEmpresa
    @IdEmpresa INT,
    @IdCentroCosto INT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @CodigoCentroCosto VARCHAR(20);

        SELECT @CodigoCentroCosto = c.Codigo
        FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
        WHERE c.IdCentroCostoConfiguracionEmpresa = @IdCentroCosto
          AND c.IdEmpresa = @IdEmpresa;

        IF @CodigoCentroCosto IS NULL
        BEGIN
            RAISERROR (N'El centro de costo no existe en la empresa activa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_AsientoDetalle AS ad
            INNER JOIN dbo.CON_Asiento AS a ON a.IdAsiento = ad.IdAsiento
            WHERE a.IdEmpresa = @IdEmpresa
              AND ad.CodigoCentroCosto = @CodigoCentroCosto
        )
        OR EXISTS
        (
            SELECT 1
            FROM dbo.BAN_MovimientoBancoDetalle AS md
            INNER JOIN dbo.BAN_MovimientoBanco AS mb ON mb.IdMovimientoBanco = md.IdMovimientoBanco
            WHERE mb.IdEmpresa = @IdEmpresa
              AND md.CodigoCentroCosto = @CodigoCentroCosto
        )
        BEGIN
            RAISERROR (N'No se puede eliminar el centro de costo porque fue utilizado en movimientos contables o bancarios.', 16, 1);
        END;

        DELETE c
        FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
        WHERE c.IdCentroCostoConfiguracionEmpresa = @IdCentroCosto
          AND c.IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH;
END;
