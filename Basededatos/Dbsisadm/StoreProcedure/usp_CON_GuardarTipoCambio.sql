-- =============================================
-- Author:        FRANCO LARA
-- Create date:   29/06/2026
-- Description:   Registra o actualiza tipos de cambio por cuenta administradora.
-- =============================================
-- Firma: FRANCO LARA - 29/06/2026 | Permite mantener tipos de cambio manuales por fecha, moneda y cuenta administradora.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarTipoCambio
    @IdTipoCambio INT = NULL,
    @IdCuentaAdministradora INT,
    @Fecha DATE,
    @IdMoneda VARCHAR(3),
    @Compra DECIMAL(18,4),
    @Venta DECIMAL(18,4),
    @CompraSBS DECIMAL(18,4),
    @VentaSBS DECIMAL(18,4),
    @Fuente VARCHAR(50),
    @UsuarioRegistro NVARCHAR(450) = NULL,
    @Estado BIT
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        SET @IdMoneda = UPPER(LTRIM(RTRIM(@IdMoneda)));
        SET @Fuente = UPPER(LTRIM(RTRIM(@Fuente)));

        IF @IdMoneda = ''
        BEGIN
            RAISERROR(N'Debe seleccionar una moneda.', 16, 1);
        END;

        IF @Fuente = ''
        BEGIN
            RAISERROR(N'Debe seleccionar una fuente.', 16, 1);
        END;

        IF @Compra <= 0 OR @Venta <= 0 OR @CompraSBS <= 0 OR @VentaSBS <= 0
        BEGIN
            RAISERROR(N'Los importes del tipo de cambio deben ser mayores a cero.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_TipoCambio AS tc
            WHERE tc.IdCuentaAdministradora = @IdCuentaAdministradora
              AND tc.Fecha = @Fecha
              AND tc.IdMoneda = @IdMoneda
              AND (@IdTipoCambio IS NULL OR tc.IdTipoCambio <> @IdTipoCambio)
        )
        BEGIN
            RAISERROR(N'Ya existe un tipo de cambio para la fecha y moneda seleccionadas.', 16, 1);
        END;

        IF @IdTipoCambio IS NULL
        BEGIN
            INSERT INTO dbo.CON_TipoCambio
            (
                IdCuentaAdministradora,
                Fecha,
                IdMoneda,
                Compra,
                Venta,
                CompraSBS,
                VentaSBS,
                Fuente,
                UsuarioRegistro,
                Estado
            )
            VALUES
            (
                @IdCuentaAdministradora,
                @Fecha,
                @IdMoneda,
                @Compra,
                @Venta,
                @CompraSBS,
                @VentaSBS,
                @Fuente,
                @UsuarioRegistro,
                @Estado
            );

            SET @IdTipoCambio = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.CON_TipoCambio
            SET Fecha = @Fecha,
                IdMoneda = @IdMoneda,
                Compra = @Compra,
                Venta = @Venta,
                CompraSBS = @CompraSBS,
                VentaSBS = @VentaSBS,
                Fuente = @Fuente,
                UsuarioRegistro = @UsuarioRegistro,
                Estado = @Estado
            WHERE IdTipoCambio = @IdTipoCambio
              AND IdCuentaAdministradora = @IdCuentaAdministradora;

            IF @@ROWCOUNT = 0
            BEGIN
                RAISERROR(N'El tipo de cambio indicado no existe para la cuenta administradora activa.', 16, 1);
            END;
        END;

        SELECT
            tc.IdTipoCambio,
            tc.IdCuentaAdministradora,
            tc.Fecha,
            tc.IdMoneda,
            tc.Compra,
            tc.Venta,
            tc.CompraSBS,
            tc.VentaSBS,
            tc.Fuente,
            tc.UsuarioRegistro,
            tc.Estado
        FROM dbo.CON_TipoCambio AS tc
        WHERE tc.IdTipoCambio = @IdTipoCambio;

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
