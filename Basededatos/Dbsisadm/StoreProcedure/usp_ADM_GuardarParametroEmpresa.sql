-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Registra o actualiza un parametro de empresa.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_ADM_GuardarParametroEmpresa
    @IdParametroEmpresa INT = NULL,
    @IdEmpresa INT,
    @TipoParametro VARCHAR(30),
    @CodigoParametro VARCHAR(100),
    @ValorParametro NVARCHAR(250),
    @DescripcionParametro NVARCHAR(300),
    @FecIni DATE = NULL,
    @FecFin DATE = NULL,
    @Activo BIT = 1,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdParametroEmpresaTrabajo INT
        DECLARE @TipoParametroTrabajo VARCHAR(30) = UPPER(LTRIM(RTRIM(@TipoParametro)))
        DECLARE @CodigoParametroTrabajo VARCHAR(100) = UPPER(LTRIM(RTRIM(@CodigoParametro)))

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.SEG_Empresa AS e
            WHERE e.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa indicada no existe.', 16, 1);
        END;

        IF @IdParametroEmpresa IS NULL
        BEGIN
            IF EXISTS
            (
                SELECT 1
                FROM dbo.ADM_ParametroEmpresa AS pe
                WHERE pe.IdEmpresa = @IdEmpresa
                  AND pe.TipoParametro = @TipoParametroTrabajo
                  AND pe.CodigoParametro = @CodigoParametroTrabajo
            )
            BEGIN
                RAISERROR(N'Ya existe un parametro con el mismo tipo y codigo para la empresa.', 16, 1);
            END;

            INSERT INTO dbo.ADM_ParametroEmpresa
            (
                IdEmpresa,
                TipoParametro,
                CodigoParametro,
                ValorParametro,
                DescripcionParametro,
                FecIni,
                FecFin,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @TipoParametroTrabajo,
                @CodigoParametroTrabajo,
                ISNULL(@ValorParametro, N''),
                ISNULL(@DescripcionParametro, N''),
                @FecIni,
                @FecFin,
                @Activo,
                @UsuarioRegistro
            );

            SET @IdParametroEmpresaTrabajo = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.ADM_ParametroEmpresa
            SET TipoParametro = @TipoParametroTrabajo,
                CodigoParametro = @CodigoParametroTrabajo,
                ValorParametro = ISNULL(@ValorParametro, N''),
                DescripcionParametro = ISNULL(@DescripcionParametro, N''),
                FecIni = @FecIni,
                FecFin = @FecFin,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND IdParametroEmpresa = @IdParametroEmpresa;

            SET @IdParametroEmpresaTrabajo = @IdParametroEmpresa;
        END;

        EXEC dbo.usp_ADM_ObtenerParametroEmpresa
            @IdEmpresa = @IdEmpresa,
            @IdParametroEmpresa = @IdParametroEmpresaTrabajo;

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
