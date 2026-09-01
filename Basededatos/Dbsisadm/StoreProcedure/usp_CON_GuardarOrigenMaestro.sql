-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Crea o actualiza un origen maestro manteniendo inmutable su codigo.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarOrigenMaestro
    @IdOrigenMaestro INT = NULL,
    @CodigoOrigen VARCHAR(10),
    @NombreOrigen NVARCHAR(150),
    @ModuloOrigen NVARCHAR(50),
    @PermiteRegistroManual BIT,
    @Estado BIT,
    @Orden INT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Codigo VARCHAR(10) = UPPER(NULLIF(LTRIM(RTRIM(@CodigoOrigen)), ''));

        IF @Codigo IS NULL
           OR NULLIF(LTRIM(RTRIM(@NombreOrigen)), N'') IS NULL
           OR NULLIF(LTRIM(RTRIM(@ModuloOrigen)), N'') IS NULL
            RAISERROR(N'El codigo, nombre y modulo del origen son obligatorios.', 16, 1);

        IF @IdOrigenMaestro IS NOT NULL
           AND NOT EXISTS (SELECT 1 FROM dbo.CON_OrigenMaestro WHERE IdOrigenMaestro = @IdOrigenMaestro)
            RAISERROR(N'El origen maestro indicado no existe.', 16, 1);

        IF @IdOrigenMaestro IS NOT NULL
           AND EXISTS
           (
               SELECT 1 FROM dbo.CON_OrigenMaestro
               WHERE IdOrigenMaestro = @IdOrigenMaestro AND CodigoOrigen <> @Codigo
           )
            RAISERROR(N'El codigo del origen no puede modificarse despues de crearlo.', 16, 1);

        IF EXISTS
        (
            SELECT 1 FROM dbo.CON_OrigenMaestro
            WHERE CodigoOrigen = @Codigo
              AND (@IdOrigenMaestro IS NULL OR IdOrigenMaestro <> @IdOrigenMaestro)
        )
            RAISERROR(N'Ya existe un origen maestro con el codigo indicado.', 16, 1);

        IF @Estado = 0
           AND EXISTS
           (
               SELECT 1
               FROM dbo.CON_ConfiguracionContabilizacionMaestro
               WHERE CodigoOrigen = @Codigo AND Activo = 1
           )
            RAISERROR(N'No se puede desactivar el origen porque esta asignado a una configuracion contable maestra activa.', 16, 1);

        IF @IdOrigenMaestro IS NULL
        BEGIN
            INSERT INTO dbo.CON_OrigenMaestro
            (
                CodigoOrigen, NombreOrigen, ModuloOrigen, PermiteRegistroManual,
                Estado, Orden, UsuarioRegistro
            )
            VALUES
            (
                @Codigo, LTRIM(RTRIM(@NombreOrigen)), UPPER(LTRIM(RTRIM(@ModuloOrigen))),
                @PermiteRegistroManual, @Estado, @Orden, @UsuarioRegistro
            );

            SET @IdOrigenMaestro = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.CON_OrigenMaestro
            SET NombreOrigen = LTRIM(RTRIM(@NombreOrigen)),
                ModuloOrigen = UPPER(LTRIM(RTRIM(@ModuloOrigen))),
                PermiteRegistroManual = @PermiteRegistroManual,
                Estado = @Estado,
                Orden = @Orden,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdOrigenMaestro = @IdOrigenMaestro;
        END;

        SELECT @IdOrigenMaestro AS IdOrigenMaestro;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
