-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   27/08/2026
-- Description:   Actualiza exclusivamente codigos contables en parametros, impuestos o documentos maestros.
-- =============================================

CREATE OR ALTER PROCEDURE dbo.usp_CON_GuardarAsignacionMaestro
    @TipoAsignacion VARCHAR(10),
    @IdRegistro INT,
    @CodigoCuenta VARCHAR(20) = NULL,
    @CodigoCuentaVentaSoles VARCHAR(20) = NULL,
    @CodigoCuentaVentaDolares VARCHAR(20) = NULL,
    @CodigoCuentaCompraSoles VARCHAR(20) = NULL,
    @CodigoCuentaCompraDolares VARCHAR(20) = NULL,
    @UsuarioRegistro NVARCHAR(450) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @Tipo VARCHAR(10) = UPPER(LTRIM(RTRIM(@TipoAsignacion)));

        DECLARE @Codigos TABLE (CodigoCuenta VARCHAR(20) NULL);
        IF @Tipo IN ('PARAMETRO', 'IMPUESTO')
            INSERT INTO @Codigos VALUES (NULLIF(LTRIM(RTRIM(@CodigoCuenta)), ''));
        ELSE IF @Tipo = 'DOCUMENTO'
            INSERT INTO @Codigos VALUES
                (NULLIF(LTRIM(RTRIM(@CodigoCuentaVentaSoles)), '')),
                (NULLIF(LTRIM(RTRIM(@CodigoCuentaVentaDolares)), '')),
                (NULLIF(LTRIM(RTRIM(@CodigoCuentaCompraSoles)), '')),
                (NULLIF(LTRIM(RTRIM(@CodigoCuentaCompraDolares)), ''));
        ELSE
            RAISERROR(N'El tipo de asignacion maestra no es valido.', 16, 1);

        IF EXISTS
        (
            SELECT 1
            FROM @Codigos AS codigo
            LEFT JOIN dbo.CON_PlanCuentaMaestro AS cuenta
                ON cuenta.CodigoCuenta = codigo.CodigoCuenta
               AND cuenta.Estado = 1
               AND cuenta.AceptaMovimiento = 1
            WHERE codigo.CodigoCuenta IS NOT NULL
              AND cuenta.IdPlanCuentaMaestro IS NULL
        )
            RAISERROR(N'Una cuenta seleccionada no existe, esta inactiva o no acepta movimiento.', 16, 1);

        IF @Tipo = 'PARAMETRO'
        BEGIN
            IF NOT EXISTS
            (
                SELECT 1 FROM dbo.ADM_ParametroMaestro
                WHERE IdParametroMaestro = @IdRegistro
                  AND CodigoParametro IN
                  (
                      'CUENTAGANANCIA', 'CUENTAGANANCIA_DC', 'CUENTAGANANCIA_AJ',
                      'CUENTAPERDIDA', 'CUENTAPERDIDA_DC', 'CUENTAPERDIDA_AJ',
                      'CTARETENCION', 'CTA_DEBE_CONSUMO', 'CTA_HABER_CONSUMO',
                      'CTADETRACCION', 'CTADEPERCEPCION'
                  )
            )
                RAISERROR(N'El parametro indicado no admite asignacion contable desde este mantenimiento.', 16, 1);

            UPDATE dbo.ADM_ParametroMaestro
            SET ValorParametro = ISNULL(NULLIF(LTRIM(RTRIM(@CodigoCuenta)), ''), N''),
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdParametroMaestro = @IdRegistro;
        END
        ELSE IF @Tipo = 'IMPUESTO'
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.CON_TipoImpuesto WHERE IdTipoImpuesto = @IdRegistro)
                RAISERROR(N'El impuesto maestro indicado no existe.', 16, 1);

            UPDATE dbo.CON_TipoImpuesto
            SET CodigoCuenta = NULLIF(LTRIM(RTRIM(@CodigoCuenta)), '')
            WHERE IdTipoImpuesto = @IdRegistro;
        END
        ELSE
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM dbo.ADM_TipoComprobante WHERE IdTipoComprobante = @IdRegistro)
                RAISERROR(N'El tipo de comprobante maestro indicado no existe.', 16, 1);

            UPDATE dbo.ADM_TipoComprobante
            SET CodigoCuentaVentaSoles = NULLIF(LTRIM(RTRIM(@CodigoCuentaVentaSoles)), ''),
                CodigoCuentaVentaDolares = NULLIF(LTRIM(RTRIM(@CodigoCuentaVentaDolares)), ''),
                CodigoCuentaCompraSoles = NULLIF(LTRIM(RTRIM(@CodigoCuentaCompraSoles)), ''),
                CodigoCuentaCompraDolares = NULLIF(LTRIM(RTRIM(@CodigoCuentaCompraDolares)), ''),
                UsuarioRegistro = @UsuarioRegistro
            WHERE IdTipoComprobante = @IdRegistro;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE()
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH
END
