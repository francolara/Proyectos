-- =============================================
-- Author:        FRANCO LARA
-- Create date:   19/04/2026
-- Firma:         Actualizacion de limites operativos del negocio desde el panel superadmin.
-- Firma:         FRANCO LARA - 18/06/2026 | Permite actualizar TipoPlan junto con los limites operativos del negocio.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Plataforma_Negocios_ActualizarLimites
    @NegocioId INT,
    @TipoPlan NVARCHAR(20),
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
        SET TipoPlan = CASE WHEN UPPER(LTRIM(RTRIM(COALESCE(@TipoPlan, N'Basico')))) = N'FULL' THEN N'Full' ELSE N'Basico' END,
            SedesPermitidas = CASE WHEN @SedesPermitidas < 1 THEN 1 ELSE @SedesPermitidas END,
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
