USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 05/05/2026 | Crea consulta de configuracion activa de proveedor de facturacion por negocio/ambiente.
CREATE OR ALTER PROCEDURE dbo.Sp_FacturacionProveedorConfig_ObtenerPorNegocio
    @NegocioId INT,
    @Ambiente NVARCHAR(15) = N'PRODUCCION'
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @AmbienteNormalizado NVARCHAR(15) = UPPER(LTRIM(RTRIM(COALESCE(@Ambiente, N'PRODUCCION'))));

        IF @AmbienteNormalizado NOT IN (N'BETA', N'PRODUCCION')
            RAISERROR('El ambiente no es valido. Valores permitidos: BETA, PRODUCCION.', 16, 1);

        SELECT
            c.Id AS NegocioProveedorConfigId,
            c.NegocioId,
            c.ProveedorId,
            p.Codigo AS ProveedorCodigo,
            p.Nombre AS ProveedorNombre,
            p.TipoAutenticacion,
            c.Ambiente,
            c.BaseUrl,
            c.ApiVersion,
            c.TimeoutSegundos,
            c.EsDefault,
            c.Activo
        FROM dbo.NegociosFacturacionProveedorConfig c
        INNER JOIN dbo.FacturacionProveedores p ON p.Id = c.ProveedorId
        WHERE c.NegocioId = @NegocioId
          AND c.Ambiente = @AmbienteNormalizado
          AND c.Activo = 1
          AND p.Activo = 1
        ORDER BY c.EsDefault DESC, c.Id ASC;

        SELECT
            cr.NegocioProveedorConfigId,
            cr.TipoCredencial,
            cr.SecretoCifrado,
            cr.KeyVersion,
            cr.ExpiraEn,
            cr.Scope
        FROM dbo.NegociosFacturacionProveedorCredencial cr
        INNER JOIN dbo.NegociosFacturacionProveedorConfig c ON c.Id = cr.NegocioProveedorConfigId
        WHERE c.NegocioId = @NegocioId
          AND c.Ambiente = @AmbienteNormalizado
          AND c.Activo = 1
          AND cr.Activo = 1;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO

