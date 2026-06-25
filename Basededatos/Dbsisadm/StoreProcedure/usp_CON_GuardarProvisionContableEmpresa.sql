-- =============================================
-- Author:        FRANCO LARA
-- Create date:   18/06/2026
-- Description:   Guarda configuracion de provision contable por modulo y empresa.
-- =============================================
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   23/06/2026
-- Description:   Amplia la configuracion de provision para compras, ventas, egresos, ingresos y aplicaciones.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarProvisionContableEmpresa
    @IdEmpresa INT,
    @ModuloOperacion VARCHAR(10),
    @IdOrigen INT,
    @GeneraAsientoAutomatico BIT,
    @Activo BIT,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN

    SET NOCOUNT ON;

    BEGIN TRY

        DECLARE @Descripcion NVARCHAR(200)

        SET @Descripcion =
            CASE @ModuloOperacion
                WHEN 'COM' THEN N'Provision Compras'
                WHEN 'VEN' THEN N'Provision Ventas'
                WHEN 'EGR' THEN N'Provision Egresos'
                WHEN 'ING' THEN N'Provision Ingresos'
                WHEN 'APNC' THEN N'Provision Aplicaciones'
                ELSE NULL
            END;

        IF @Descripcion IS NULL
        BEGIN
            RAISERROR(N'El modulo de provision es invalido.', 16, 1);
        END;

        IF NOT EXISTS
        (
            SELECT 1
            FROM dbo.CON_Origen AS o
            WHERE o.IdOrigen = @IdOrigen
              AND o.IdEmpresa = @IdEmpresa
              AND o.Estado = 1
        )
        BEGIN
            RAISERROR(N'El origen indicado no existe o no pertenece a la empresa.', 16, 1);
        END;

        IF EXISTS
        (
            SELECT 1
            FROM dbo.CON_ConfiguracionContabilizacion AS c
            WHERE c.IdEmpresa = @IdEmpresa
              AND c.ModuloOperacion = @ModuloOperacion
              AND c.EscenarioOperacion = 'PROVISION'
        )
        BEGIN
            UPDATE dbo.CON_ConfiguracionContabilizacion
            SET IdOrigen = @IdOrigen,
                Descripcion = @Descripcion,
                GeneraAsientoAutomatico = @GeneraAsientoAutomatico,
                UsaTipoCambio = 1,
                Activo = @Activo,
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdEmpresa = @IdEmpresa
              AND ModuloOperacion = @ModuloOperacion
              AND EscenarioOperacion = 'PROVISION';
        END
        ELSE
        BEGIN
            INSERT INTO dbo.CON_ConfiguracionContabilizacion
            (
                IdEmpresa,
                ModuloOperacion,
                EscenarioOperacion,
                IdOrigen,
                Descripcion,
                GeneraAsientoAutomatico,
                UsaTipoCambio,
                Activo,
                UsuarioRegistro
            )
            VALUES
            (
                @IdEmpresa,
                @ModuloOperacion,
                'PROVISION',
                @IdOrigen,
                @Descripcion,
                @GeneraAsientoAutomatico,
                1,
                @Activo,
                @UsuarioRegistro
            );
        END;

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
