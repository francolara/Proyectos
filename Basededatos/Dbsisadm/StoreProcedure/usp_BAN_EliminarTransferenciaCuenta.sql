-- =============================================
-- Author:        FRANCO LARA
-- Create date:   24/06/2026
-- Description:   Elimina una transferencia entre cuentas borrando ambos movimientos bancarios enlazados y limpiando primero las referencias a IdAsiento para evitar conflictos de llave foranea.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_BAN_EliminarTransferenciaCuenta
    @IdEmpresa INT,
    @IdMovimientoBancoEmisor INT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdTransferenciaCuenta UNIQUEIDENTIFIER = NULL;
        DECLARE @Asientos TABLE
        (
            IdAsiento INT NOT NULL PRIMARY KEY
        );

        SELECT
            @IdTransferenciaCuenta = m.IdTransferenciaCuenta
        FROM dbo.BAN_MovimientoBanco AS m
        WHERE m.IdMovimientoBanco = @IdMovimientoBancoEmisor
          AND m.IdEmpresa = @IdEmpresa
          AND m.RolTransferencia = 'E'
          AND m.Activo = 1;

        IF @IdTransferenciaCuenta IS NULL
        BEGIN
            RAISERROR('La transferencia seleccionada no existe para la empresa activa.', 16, 1);
        END;

        BEGIN TRANSACTION;

        INSERT INTO @Asientos (IdAsiento)
        SELECT DISTINCT m.IdAsiento
        FROM dbo.BAN_MovimientoBanco AS m
        WHERE m.IdEmpresa = @IdEmpresa
          AND m.IdTransferenciaCuenta = @IdTransferenciaCuenta
          AND m.IdAsiento IS NOT NULL;

        DELETE ad
        FROM dbo.CON_AsientoDetalle AS ad
        INNER JOIN @Asientos AS a
            ON a.IdAsiento = ad.IdAsiento;

        DELETE d
        FROM dbo.BAN_MovimientoBancoDetalle AS d
        INNER JOIN dbo.BAN_MovimientoBanco AS m
            ON m.IdMovimientoBanco = d.IdMovimientoBanco
        WHERE m.IdEmpresa = @IdEmpresa
          AND m.IdTransferenciaCuenta = @IdTransferenciaCuenta;

        UPDATE dbo.BAN_MovimientoBanco
        SET IdMovimientoBancoRelacionado = NULL,
            IdAsiento = NULL
        WHERE IdEmpresa = @IdEmpresa
          AND IdTransferenciaCuenta = @IdTransferenciaCuenta;

        DELETE FROM dbo.BAN_MovimientoBanco
        WHERE IdEmpresa = @IdEmpresa
          AND IdTransferenciaCuenta = @IdTransferenciaCuenta;

        DELETE a
        FROM dbo.CON_Asiento AS a
        INNER JOIN @Asientos AS x
            ON x.IdAsiento = a.IdAsiento;

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
        BEGIN
            ROLLBACK TRANSACTION;
        END;

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
