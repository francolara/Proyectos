-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Inserta o actualiza un centro de costo configurado por empresa validando codigo unico.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarCentroCostoConfiguracionEmpresa
    @IdCentroCosto INT = NULL,
    @IdEmpresa INT,
    @CodigoCentroCosto VARCHAR(20),
    @NombreCentroCosto NVARCHAR(150),
    @Estado BIT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.Codigo = @CodigoCentroCosto
              AND (@IdCentroCosto IS NULL OR c.IdCentroCostoConfiguracionEmpresa <> @IdCentroCosto)
        )
        BEGIN
            RAISERROR(N'Ya existe un centro de costo con el mismo codigo para la empresa activa.', 16, 1);
        END;

        IF @IdCentroCosto IS NULL
        BEGIN
            INSERT INTO dbo.CON_CentroCostoConfiguracionEmpresa
            (
                IdEmpresa,
                Codigo,
                Nombre,
                Estado
            )
            VALUES
            (
                @IdEmpresa,
                @CodigoCentroCosto,
                @NombreCentroCosto,
                @Estado
            );

            SET @IdCentroCosto = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
                WHERE c.IdCentroCostoConfiguracionEmpresa = @IdCentroCosto
                  AND c.IdEmpresa = @IdEmpresa
            )
            BEGIN
                RAISERROR(N'El centro de costo a actualizar no existe en la empresa activa.', 16, 1);
            END;

            UPDATE dbo.CON_CentroCostoConfiguracionEmpresa
            SET Codigo = @CodigoCentroCosto,
                Nombre = @NombreCentroCosto,
                Estado = @Estado
            WHERE IdCentroCostoConfiguracionEmpresa = @IdCentroCosto
              AND IdEmpresa = @IdEmpresa;
        END;

        SELECT
            c.IdCentroCostoConfiguracionEmpresa AS IdCentroCosto,
            c.IdEmpresa,
            c.Codigo AS CodigoCentroCosto,
            c.Nombre AS NombreCentroCosto,
            c.Estado
        FROM dbo.CON_CentroCostoConfiguracionEmpresa AS c
        WHERE c.IdCentroCostoConfiguracionEmpresa = @IdCentroCosto;

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
