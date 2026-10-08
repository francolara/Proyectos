
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: FRANCO LARA - 01/10/2026 | Entrega el siguiente correlativo visible de una entidad por negocio, con bloqueo transaccional para impedir duplicados concurrentes.
CREATE OR ALTER PROCEDURE dbo.Sp_NegocioCorrelativos_ObtenerSiguiente
    @NegocioId INT,
    @Entidad NVARCHAR(30),
    @Usuario NVARCHAR(200) = NULL,
    @Numero INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Entidad = UPPER(LTRIM(RTRIM(@Entidad)));

        IF @Entidad NOT IN (N'RESERVA', N'PAGO', N'CLIENTE', N'SEDE', N'ESPACIO', N'CUPON', N'PROMOCION')
            RAISERROR('La entidad no admite correlativo por negocio.', 16, 1);

        BEGIN TRANSACTION;

        UPDATE dbo.NegocioCorrelativos WITH (UPDLOCK, SERIALIZABLE)
        SET UltimoNumero = UltimoNumero + 1,
            FechaActualizacion = SYSUTCDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE NegocioId = @NegocioId
          AND Entidad = @Entidad;

        IF @@ROWCOUNT = 0
        BEGIN
            INSERT INTO dbo.NegocioCorrelativos
            (
                NegocioId, Entidad, UltimoNumero, FechaActualizacion, UsuarioActualizacion
            )
            VALUES
            (
                @NegocioId, @Entidad, 1, SYSUTCDATETIME(), @Usuario
            );

            SET @Numero = 1;
        END
        ELSE
        BEGIN
            SELECT @Numero = UltimoNumero
            FROM dbo.NegocioCorrelativos WITH (UPDLOCK, HOLDLOCK)
            WHERE NegocioId = @NegocioId
              AND Entidad = @Entidad;
        END

        COMMIT TRANSACTION;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0
            ROLLBACK TRANSACTION;

        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
