
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 10/06/2026 | Activa o desactiva boletines deportivos desde administracion sin eliminar el historial.
-- Firma: Codex - 11/06/2026 | Devuelve el estado real persistido del boletin despues del cambio para validar activacion o desactivacion en el panel.
CREATE OR ALTER PROCEDURE dbo.Sp_BoletinesDeportivos_CambiarEstado
    @IdBoletin INT,
    @Activo BIT,
    @Usuario NVARCHAR(120)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        UPDATE dbo.BoletinesDeportivos
        SET Activo = @Activo,
            FechaActualizacion = SYSDATETIME(),
            UsuarioActualizacion = @Usuario
        WHERE IdBoletin = @IdBoletin;

        IF @@ROWCOUNT = 0
            RAISERROR(N'No se encontro el boletin deportivo a actualizar.', 16, 1);

        SELECT TOP (1)
            Activo
        FROM dbo.BoletinesDeportivos
        WHERE IdBoletin = @IdBoletin;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
