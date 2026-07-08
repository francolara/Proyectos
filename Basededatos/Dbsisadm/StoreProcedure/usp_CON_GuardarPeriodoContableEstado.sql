-- =============================================
-- Author:        FRANCO LARA
-- Create date:   30/06/2026
-- Description:   Cierra o reabre un periodo contable por empresa.
-- =============================================
-- Firma: FRANCO LARA - 30/06/2026 | Permite cerrar o abrir periodos contables para bloquear la operativa de compras, ventas, bancos, transferencias y aplicaciones.

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarPeriodoContableEstado
    @IdEmpresa INT,
    @Periodo CHAR(6),
    @Cerrado BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @IdPeriodoContableEstado INT

        SET @Periodo = LTRIM(RTRIM(@Periodo));

        IF @Periodo NOT LIKE '[1-2][0-9][0-9][0-9][0-1][0-9]'
        BEGIN
            RAISERROR(N'El periodo debe tener formato YYYYMM.', 16, 1);
        END;

        SELECT
            @IdPeriodoContableEstado = pce.IdPeriodoContableEstado
        FROM dbo.CON_PeriodoContableEstado AS pce
        WHERE pce.IdEmpresa = @IdEmpresa
          AND pce.Periodo = @Periodo;

        IF @IdPeriodoContableEstado IS NULL
        BEGIN
            INSERT INTO dbo.CON_PeriodoContableEstado
            (
                IdEmpresa,
                Periodo,
                Cerrado,
                UsuarioRegistro,
                FechaCierre,
                UsuarioCierre,
                FechaApertura,
                UsuarioApertura
            )
            VALUES
            (
                @IdEmpresa,
                @Periodo,
                @Cerrado,
                @UsuarioRegistro,
                CASE WHEN @Cerrado = 1 THEN SYSDATETIME() ELSE NULL END,
                CASE WHEN @Cerrado = 1 THEN @UsuarioRegistro ELSE NULL END,
                CASE WHEN @Cerrado = 0 THEN SYSDATETIME() ELSE NULL END,
                CASE WHEN @Cerrado = 0 THEN @UsuarioRegistro ELSE NULL END
            );

            SET @IdPeriodoContableEstado = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.CON_PeriodoContableEstado
            SET Cerrado = @Cerrado,
                UsuarioRegistro = @UsuarioRegistro,
                FechaCierre = CASE
                    WHEN @Cerrado = 1 THEN SYSDATETIME()
                    ELSE FechaCierre
                END,
                UsuarioCierre = CASE
                    WHEN @Cerrado = 1 THEN @UsuarioRegistro
                    ELSE UsuarioCierre
                END,
                FechaApertura = CASE
                    WHEN @Cerrado = 0 THEN SYSDATETIME()
                    ELSE FechaApertura
                END,
                UsuarioApertura = CASE
                    WHEN @Cerrado = 0 THEN @UsuarioRegistro
                    ELSE UsuarioApertura
                END
            WHERE IdPeriodoContableEstado = @IdPeriodoContableEstado;
        END;

        SELECT
            pce.IdPeriodoContableEstado,
            pce.IdEmpresa,
            pce.Periodo,
            pce.Cerrado,
            pce.FechaRegistro,
            pce.UsuarioRegistro,
            pce.FechaCierre,
            pce.UsuarioCierre,
            pce.FechaApertura,
            pce.UsuarioApertura
        FROM dbo.CON_PeriodoContableEstado AS pce
        WHERE pce.IdPeriodoContableEstado = @IdPeriodoContableEstado;

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
