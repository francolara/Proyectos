-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   07/07/2026
-- Description:   Registra el metadato de una generación PLE sin almacenar el contenido del archivo TXT.
-- =============================================

-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   04/08/2026
-- Description:   Registra por separado el complemento y conserva el snapshot del plan sin marcarlo como presentado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_PLE_Historial_Registrar
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @CodigoLibro VARCHAR(10),
    @CodigoFormato VARCHAR(10),
    @NombreArchivo NVARCHAR(250),
    @CantidadRegistros INT,
    @TotalDebe DECIMAL(18,2),
    @TotalHaber DECIMAL(18,2),
    @Estado NVARCHAR(20),
    @Observaciones NVARCHAR(MAX) = NULL,
    @UsuarioGeneracion NVARCHAR(450) = NULL,
    @CodigoFormatoComplementario VARCHAR(10) = NULL,
    @NombreArchivoComplementario NVARCHAR(250) = NULL,
    @CantidadRegistrosComplementario INT = 0,
    @HuellaPlanContable CHAR(64) = NULL,
    @PlanContableSnapshot NVARCHAR(MAX) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        INSERT INTO dbo.CON_LibroElectronicoGeneracion
        (
            IdEmpresa,
            Periodo,
            CodigoLibro,
            CodigoFormato,
            NombreArchivo,
            CantidadRegistros,
            TotalDebe,
            TotalHaber,
            Estado,
            Observaciones,
            FechaGeneracion,
            UsuarioGeneracion,
            CodigoFormatoComplementario,
            NombreArchivoComplementario,
            CantidadRegistrosComplementario,
            HuellaPlanContable,
            PlanContableSnapshot,
            PlanPresentado
        )
        VALUES
        (
            @IdEmpresa,
            @Periodo,
            @CodigoLibro,
            @CodigoFormato,
            @NombreArchivo,
            @CantidadRegistros,
            @TotalDebe,
            @TotalHaber,
            @Estado,
            @Observaciones,
            SYSDATETIME(),
            @UsuarioGeneracion,
            @CodigoFormatoComplementario,
            @NombreArchivoComplementario,
            @CantidadRegistrosComplementario,
            @HuellaPlanContable,
            @PlanContableSnapshot,
            0
        );

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
