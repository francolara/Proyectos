-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Copia plan de cuentas maestro interno hacia una empresa con ColBalance, moneda y tipo de cambio.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   02/07/2026
-- Description:   Permite cargar el plan desde una empresa base o desde el maestro, heredando diferencia en cambio por analisis.
-- =============================================
-- Firma: FRANCO LARA - 25/08/2026 | El maestro no contiene GeneraDiferenciaPorAnalisis; al cargarlo se inicializa en cero para la empresa.

CREATE OR ALTER PROCEDURE dbo.usp_CON_CargarPlanCuentaDefaultEmpresa
    @IdEmpresa INT,
    @IdEmpresaBase INT = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_PlanCuenta AS pc
            WHERE pc.IdEmpresa = @IdEmpresa
        )
        BEGIN
            RAISERROR(N'La empresa ya tiene plan de cuentas registrado.', 16, 1);
        END;

        IF @IdEmpresaBase IS NOT NULL AND @IdEmpresaBase = @IdEmpresa
        BEGIN
            RAISERROR(N'La empresa base debe ser distinta de la empresa destino.', 16, 1);
        END;

        IF @IdEmpresaBase IS NOT NULL
           AND NOT EXISTS
           (
               SELECT 1
               FROM dbo.SEG_Empresa AS e
               WHERE e.IdEmpresa = @IdEmpresaBase
           )
        BEGIN
            RAISERROR(N'La empresa base indicada no existe.', 16, 1);
        END;

        CREATE TABLE #PlanCuentaFuente
        (
            CodigoCuenta VARCHAR(20) NOT NULL,
            CodigoCuentaPadre VARCHAR(20) NULL,
            NombreCuenta NVARCHAR(200) NOT NULL,
            NivelCuenta TINYINT NOT NULL,
            ColBalance CHAR(1) NOT NULL,
            IdMoneda VARCHAR(3) NOT NULL,
            TipoCambio CHAR(1) NOT NULL,
            AceptaMovimiento BIT NOT NULL,
            GeneraDiferenciaPorAnalisis BIT NOT NULL,
            RequiereCentroCosto BIT NOT NULL,
            Estado BIT NOT NULL,
            Orden INT NOT NULL
        );

        IF @IdEmpresaBase IS NOT NULL
           AND EXISTS
           (
               SELECT 1
               FROM dbo.CON_PlanCuenta AS pc
               WHERE pc.IdEmpresa = @IdEmpresaBase
           )
        BEGIN
            INSERT INTO #PlanCuentaFuente
            (
                CodigoCuenta,
                CodigoCuentaPadre,
                NombreCuenta,
                NivelCuenta,
                ColBalance,
                IdMoneda,
                TipoCambio,
                AceptaMovimiento,
                GeneraDiferenciaPorAnalisis,
                RequiereCentroCosto,
                Estado,
                Orden
            )
            SELECT
                hijo.CodigoCuenta,
                padre.CodigoCuenta,
                hijo.NombreCuenta,
                hijo.NivelCuenta,
                hijo.ColBalance,
                hijo.IdMoneda,
                hijo.TipoCambio,
                hijo.AceptaMovimiento,
                hijo.GeneraDiferenciaPorAnalisis,
                hijo.RequiereCentroCosto,
                hijo.Estado,
                hijo.Orden
            FROM dbo.CON_PlanCuenta AS hijo
            LEFT JOIN dbo.CON_PlanCuenta AS padre
                ON padre.IdPlanCuenta = hijo.IdPlanCuentaPadre
            WHERE hijo.IdEmpresa = @IdEmpresaBase
              AND hijo.Estado = 1;
        END
        ELSE
        BEGIN
            INSERT INTO #PlanCuentaFuente
            (
                CodigoCuenta,
                CodigoCuentaPadre,
                NombreCuenta,
                NivelCuenta,
                ColBalance,
                IdMoneda,
                TipoCambio,
                AceptaMovimiento,
                GeneraDiferenciaPorAnalisis,
                RequiereCentroCosto,
                Estado,
                Orden
            )
            SELECT
                pcm.CodigoCuenta,
                pcm.CodigoCuentaPadre,
                pcm.NombreCuenta,
                pcm.NivelCuenta,
                pcm.ColBalance,
                pcm.IdMoneda,
                pcm.TipoCambio,
                pcm.AceptaMovimiento,
                CAST(0 AS BIT) AS GeneraDiferenciaPorAnalisis,
                pcm.RequiereCentroCosto,
                pcm.Estado,
                pcm.Orden
            FROM dbo.CON_PlanCuentaMaestro AS pcm
            WHERE pcm.Estado = 1;
        END;

        INSERT INTO dbo.CON_PlanCuenta
        (
            IdEmpresa,
            IdPlanCuentaPadre,
            CodigoCuenta,
            NombreCuenta,
            NivelCuenta,
            ColBalance,
            IdMoneda,
            TipoCambio,
            AceptaMovimiento,
            GeneraDiferenciaPorAnalisis,
            RequiereCentroCosto,
            Estado,
            UsuarioRegistro
        )
        SELECT
            @IdEmpresa,
            NULL,
            fuente.CodigoCuenta,
            fuente.NombreCuenta,
            fuente.NivelCuenta,
            fuente.ColBalance,
            fuente.IdMoneda,
            fuente.TipoCambio,
            fuente.AceptaMovimiento,
            fuente.GeneraDiferenciaPorAnalisis,
            fuente.RequiereCentroCosto,
            fuente.Estado,
            @UsuarioRegistro
        FROM #PlanCuentaFuente AS fuente
        WHERE fuente.Estado = 1
        ORDER BY fuente.NivelCuenta, fuente.Orden, fuente.CodigoCuenta;

        UPDATE hijo
        SET IdPlanCuentaPadre = padre.IdPlanCuenta
        FROM dbo.CON_PlanCuenta AS hijo
        INNER JOIN #PlanCuentaFuente AS fuenteHijo
            ON fuenteHijo.CodigoCuenta = hijo.CodigoCuenta
        INNER JOIN dbo.CON_PlanCuenta AS padre
            ON padre.IdEmpresa = hijo.IdEmpresa
           AND padre.CodigoCuenta = fuenteHijo.CodigoCuentaPadre
        WHERE hijo.IdEmpresa = @IdEmpresa;

        DROP TABLE #PlanCuentaFuente;

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
