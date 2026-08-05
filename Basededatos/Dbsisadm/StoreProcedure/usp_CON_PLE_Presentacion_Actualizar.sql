-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Marca o desmarca una generacion como presentada sin romper la continuidad posterior.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_Presentacion_Actualizar
    @IdEmpresa INT,
    @IdLibroElectronicoGeneracion INT,
    @Presentado BIT,
    @UsuarioPresentacion NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Periodo CHAR(6);
        DECLARE @CodigoLibro VARCHAR(10);

        SELECT @Periodo = h.Periodo, @CodigoLibro = h.CodigoLibro
        FROM dbo.CON_LibroElectronicoGeneracion AS h
        WHERE h.IdLibroElectronicoGeneracion = @IdLibroElectronicoGeneracion
          AND h.IdEmpresa = @IdEmpresa;

        IF @Periodo IS NULL
            RAISERROR (N'No se encontro la generacion seleccionada para la empresa activa.', 16, 1);

        IF @IdLibroElectronicoGeneracion <>
        (
            SELECT MAX(h.IdLibroElectronicoGeneracion)
            FROM dbo.CON_LibroElectronicoGeneracion AS h
            WHERE h.IdEmpresa = @IdEmpresa
              AND h.Periodo = @Periodo
              AND h.CodigoLibro = @CodigoLibro
        )
            RAISERROR (N'Solo se puede confirmar la generacion mas reciente del periodo.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_LibroElectronicoGeneracion AS h
            WHERE h.IdEmpresa = @IdEmpresa
              AND h.CodigoLibro = @CodigoLibro
              AND h.Periodo > @Periodo
              AND h.PlanPresentado = 1
        )
            RAISERROR (N'No se puede cambiar el periodo porque existe una presentacion posterior.', 16, 1);

        IF @Presentado = 1
        BEGIN
            UPDATE dbo.CON_LibroElectronicoGeneracion
            SET PlanPresentado = 0,
                FechaPresentacion = NULL,
                UsuarioPresentacion = NULL
            WHERE IdEmpresa = @IdEmpresa
              AND Periodo = @Periodo
              AND CodigoLibro = @CodigoLibro;
        END;

        UPDATE dbo.CON_LibroElectronicoGeneracion
        SET PlanPresentado = @Presentado,
            FechaPresentacion = CASE WHEN @Presentado = 1 THEN SYSDATETIME() ELSE NULL END,
            UsuarioPresentacion = CASE WHEN @Presentado = 1 THEN @UsuarioPresentacion ELSE NULL END
        WHERE IdLibroElectronicoGeneracion = @IdLibroElectronicoGeneracion
          AND IdEmpresa = @IdEmpresa;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
