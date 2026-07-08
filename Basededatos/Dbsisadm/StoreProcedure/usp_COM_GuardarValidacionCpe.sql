-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Guarda la fecha, estado y mensaje de la ultima validacion CPE ejecutada sobre una compra.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_COM_GuardarValidacionCpe
    @IdCompra INT,
    @IdEmpresa INT,
    @FechaValidacionCpe DATETIME2(0),
    @EstadoValidacionCpe NVARCHAR(50),
    @MensajeValidacionCpe NVARCHAR(500) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.COM_Compra AS c
            WHERE c.IdCompra = @IdCompra
              AND c.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La compra seleccionada no existe en la empresa activa.', 16, 1);
        END;

        UPDATE dbo.COM_Compra
        SET FechaValidacionCpe = @FechaValidacionCpe,
            EstadoValidacionCpe = NULLIF(LTRIM(RTRIM(@EstadoValidacionCpe)), N''),
            MensajeValidacionCpe = NULLIF(LTRIM(RTRIM(@MensajeValidacionCpe)), N''),
            UsuarioRegistro = @UsuarioRegistro
        WHERE IdCompra = @IdCompra
          AND IdEmpresa = @IdEmpresa;

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
