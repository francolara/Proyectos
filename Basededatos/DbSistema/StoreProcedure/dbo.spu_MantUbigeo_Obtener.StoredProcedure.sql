-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene un país o registro de ubigeo para su mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUbigeo_Obtener
    @Tipo TINYINT,
    @Codigo VARCHAR(6),
    @IdPais VARCHAR(5) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF @Tipo = 0
        BEGIN
            SELECT D.GlsDato
            FROM dbo.Datos AS D
            WHERE D.IdDato = @Codigo AND D.IdTipoDatos = '02';
            RETURN;
        END;

        SELECT U.GlsUbigeo, U.IdZona
        FROM dbo.Ubigeo AS U
        WHERE U.IdDistrito = @Codigo AND U.IdPais = @IdPais;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
