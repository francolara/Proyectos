-- =============================================
-- Author:        FRANCO LARA
-- Create date:   16/06/2026
-- Description:   Inserta o actualiza un origen contable por empresa validando codigo y estado.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarOrigenPorEmpresa
    @IdOrigen INT = NULL,
    @IdEmpresa INT,
    @CodigoOrigen VARCHAR(10),
    @NombreOrigen NVARCHAR(150),
    @ModuloOrigen NVARCHAR(50),
    @PermiteRegistroManual BIT,
    @Estado BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
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
            FROM dbo.CON_Origen AS o
            WHERE o.IdEmpresa = @IdEmpresa
              AND o.CodigoOrigen = @CodigoOrigen
              AND (@IdOrigen IS NULL OR o.IdOrigen <> @IdOrigen)
        )
        BEGIN
            RAISERROR(N'Ya existe un origen con el mismo codigo para la empresa activa.', 16, 1);
        END;

        IF @IdOrigen IS NULL
        BEGIN
            INSERT INTO dbo.CON_Origen
            (
                IdEmpresa,
                CodigoOrigen,
                NombreOrigen,
                ModuloOrigen,
                PermiteRegistroManual,
                Estado,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @CodigoOrigen,
                @NombreOrigen,
                @ModuloOrigen,
                @PermiteRegistroManual,
                @Estado,
                @UsuarioRegistro
            );

            SET @IdOrigen = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1
                FROM dbo.CON_Origen AS o
                WHERE o.IdOrigen = @IdOrigen
                  AND o.IdEmpresa = @IdEmpresa
            )
            BEGIN
                RAISERROR(N'El origen a actualizar no existe en la empresa activa.', 16, 1);
            END;

            UPDATE dbo.CON_Origen
            SET CodigoOrigen = @CodigoOrigen,
                NombreOrigen = @NombreOrigen,
                ModuloOrigen = @ModuloOrigen,
                PermiteRegistroManual = @PermiteRegistroManual,
                Estado = @Estado,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdOrigen = @IdOrigen
              AND IdEmpresa = @IdEmpresa;
        END;

        SELECT
            o.IdOrigen,
            o.CodigoOrigen,
            o.NombreOrigen,
            o.ModuloOrigen,
            o.PermiteRegistroManual,
            o.Estado
        FROM dbo.CON_Origen AS o
        WHERE o.IdOrigen = @IdOrigen;

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
