USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 05/05/2026 | Crea alta/actualizacion de configuracion de proveedor de facturacion por negocio.
CREATE OR ALTER PROCEDURE dbo.Sp_FacturacionProveedorConfig_Guardar
    @NegocioProveedorConfigId INT = NULL,
    @NegocioId INT,
    @ProveedorId INT,
    @Ambiente NVARCHAR(15),
    @BaseUrl NVARCHAR(500),
    @ApiVersion NVARCHAR(20) = NULL,
    @TimeoutSegundos INT = 30,
    @EsDefault BIT = 0,
    @Activo BIT = 1,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @AmbienteNormalizado NVARCHAR(15) = UPPER(LTRIM(RTRIM(COALESCE(@Ambiente, N''))));
        DECLARE @BaseUrlNormalizada NVARCHAR(500) = NULLIF(LTRIM(RTRIM(@BaseUrl)), N'');
        DECLARE @ApiVersionNormalizada NVARCHAR(20) = NULLIF(LTRIM(RTRIM(@ApiVersion)), N'');

        IF NOT EXISTS (SELECT 1 FROM dbo.Negocios WHERE Id = @NegocioId AND Activo = 1)
            RAISERROR('El negocio no existe o esta inactivo.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.FacturacionProveedores WHERE Id = @ProveedorId AND Activo = 1)
            RAISERROR('El proveedor de facturacion no existe o esta inactivo.', 16, 1);

        IF @AmbienteNormalizado NOT IN (N'BETA', N'PRODUCCION')
            RAISERROR('El ambiente no es valido. Valores permitidos: BETA, PRODUCCION.', 16, 1);

        IF @BaseUrlNormalizada IS NULL
            RAISERROR('La URL base del proveedor es obligatoria.', 16, 1);

        IF @TimeoutSegundos < 5 OR @TimeoutSegundos > 300
            RAISERROR('El timeout debe estar entre 5 y 300 segundos.', 16, 1);

        IF @NegocioProveedorConfigId IS NULL
        BEGIN
            IF EXISTS (
                SELECT 1
                FROM dbo.NegociosFacturacionProveedorConfig
                WHERE NegocioId = @NegocioId
                  AND ProveedorId = @ProveedorId
                  AND Ambiente = @AmbienteNormalizado
            )
                RAISERROR('Ya existe una configuracion para el negocio/proveedor/ambiente.', 16, 1);

            INSERT INTO dbo.NegociosFacturacionProveedorConfig
            (
                NegocioId,
                ProveedorId,
                Ambiente,
                BaseUrl,
                ApiVersion,
                TimeoutSegundos,
                EsDefault,
                Activo,
                FechaRegistro,
                FechaActualizacion,
                UsuarioActualizacion
            )
            VALUES
            (
                @NegocioId,
                @ProveedorId,
                @AmbienteNormalizado,
                @BaseUrlNormalizada,
                @ApiVersionNormalizada,
                @TimeoutSegundos,
                @EsDefault,
                @Activo,
                SYSUTCDATETIME(),
                SYSUTCDATETIME(),
                @Usuario
            );

            SET @NegocioProveedorConfigId = SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE c
            SET
                c.ProveedorId = @ProveedorId,
                c.Ambiente = @AmbienteNormalizado,
                c.BaseUrl = @BaseUrlNormalizada,
                c.ApiVersion = @ApiVersionNormalizada,
                c.TimeoutSegundos = @TimeoutSegundos,
                c.EsDefault = @EsDefault,
                c.Activo = @Activo,
                c.FechaActualizacion = SYSUTCDATETIME(),
                c.UsuarioActualizacion = @Usuario
            FROM dbo.NegociosFacturacionProveedorConfig c
            WHERE c.Id = @NegocioProveedorConfigId
              AND c.NegocioId = @NegocioId;

            IF @@ROWCOUNT = 0
                RAISERROR('No se encontro la configuracion de proveedor para actualizar.', 16, 1);
        END

        IF @EsDefault = 1
        BEGIN
            UPDATE c
            SET
                c.EsDefault = CASE WHEN c.Id = @NegocioProveedorConfigId THEN 1 ELSE 0 END,
                c.FechaActualizacion = SYSUTCDATETIME(),
                c.UsuarioActualizacion = @Usuario
            FROM dbo.NegociosFacturacionProveedorConfig c
            WHERE c.NegocioId = @NegocioId
              AND c.Ambiente = @AmbienteNormalizado
              AND c.Activo = 1;

            UPDATE n
            SET
                n.ProveedorElectronicoDefaultId = @ProveedorId
            FROM dbo.Negocios n
            WHERE n.Id = @NegocioId;
        END

        SELECT
            c.Id,
            c.NegocioId,
            c.ProveedorId,
            c.Ambiente,
            c.BaseUrl,
            c.ApiVersion,
            c.TimeoutSegundos,
            c.EsDefault,
            c.Activo
        FROM dbo.NegociosFacturacionProveedorConfig c
        WHERE c.Id = @NegocioProveedorConfigId;
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

