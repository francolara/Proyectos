-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/11/2026
-- Description:   Guarda poliza SCTR y trabajadores cubiertos en una sola transaccion.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_PolizaSctrHYT_Guardar
    @UsuarioId NVARCHAR(128),
    @Aseguradora NVARCHAR(120),
    @NumeroPoliza NVARCHAR(80),
    @Cobertura VARCHAR(10),
    @VigenciaDesde DATE,
    @VigenciaHasta DATE,
    @PolizaId BIGINT = NULL,
    @VersionFila BINARY(8) = NULL,
    @Trabajadores NVARCHAR(MAX) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId, N'OBLIGACIONES.EDITAR') = 0
            THROW 51000, N'No tiene permiso para guardar polizas SCTR.', 1;

        SET @Aseguradora = NULLIF(TRIM(@Aseguradora), N'');
        SET @NumeroPoliza = NULLIF(TRIM(@NumeroPoliza), N'');
        IF @Aseguradora IS NULL OR @NumeroPoliza IS NULL OR @Cobertura IS NULL
           OR @Cobertura NOT IN ('SALUD', 'PENSION') OR @VigenciaDesde IS NULL
           OR @VigenciaHasta IS NULL OR @VigenciaHasta < @VigenciaDesde
            THROW 51000, N'Revise la aseguradora, la cobertura y la vigencia de la poliza.', 1;

        DECLARE @Personal TABLE (TrabajadorId BIGINT PRIMARY KEY);
        IF @Trabajadores IS NOT NULL
        BEGIN
            IF COALESCE(ISJSON(@Trabajadores, ARRAY), 0) <> 1
                THROW 51000, N'Seleccione trabajadores validos.', 1;

            INSERT INTO @Personal (TrabajadorId)
            SELECT DISTINCT TRY_CONVERT(BIGINT, fila.value)
            FROM OPENJSON(@Trabajadores) AS fila
            WHERE TRY_CONVERT(BIGINT, fila.value) > 0;
            IF NOT EXISTS (SELECT 1 FROM @Personal)
                THROW 51000, N'Seleccione al menos un trabajador.', 1;
        END;

        BEGIN TRANSACTION;
        DECLARE @Accion NVARCHAR(60) = N'MODIFICAR';
        IF @PolizaId IS NULL
        BEGIN
            INSERT INTO dbo.PolizaSctrHYT (Aseguradora, NumeroPoliza, Cobertura, VigenciaDesde, VigenciaHasta)
            VALUES (@Aseguradora, @NumeroPoliza, @Cobertura, @VigenciaDesde, @VigenciaHasta);
            SET @PolizaId = SCOPE_IDENTITY();
            SET @Accion = N'REGISTRAR';
        END
        ELSE
        BEGIN
            UPDATE dbo.PolizaSctrHYT
            SET Aseguradora = @Aseguradora, NumeroPoliza = @NumeroPoliza, Cobertura = @Cobertura,
                VigenciaDesde = @VigenciaDesde, VigenciaHasta = @VigenciaHasta
            WHERE PolizaId = @PolizaId AND VersionFila = @VersionFila;
            IF @@ROWCOUNT = 0
                THROW 51000, N'La poliza cambio o no existe. Vuelva a abrirla.', 1;
        END;

        IF @Trabajadores IS NOT NULL
        BEGIN
            IF EXISTS (
                SELECT 1 FROM @Personal AS seleccionado
                LEFT JOIN dbo.TrabajadorHYT AS trabajador WITH (HOLDLOCK) ON trabajador.TrabajadorId = seleccionado.TrabajadorId
                LEFT JOIN dbo.PolizaSctrTrabajadorHYT AS existente
                    ON existente.PolizaId = @PolizaId AND existente.TrabajadorId = seleccionado.TrabajadorId
                WHERE trabajador.TrabajadorId IS NULL OR (trabajador.Activo = 0 AND existente.TrabajadorId IS NULL))
                THROW 51000, N'Uno de los trabajadores seleccionados no esta disponible.', 1;

            DELETE FROM dbo.PolizaSctrTrabajadorHYT WHERE PolizaId = @PolizaId;
            INSERT INTO dbo.PolizaSctrTrabajadorHYT (PolizaId, TrabajadorId)
            SELECT @PolizaId, seleccionado.TrabajadorId FROM @Personal AS seleccionado;
        END;

        DECLARE @RegistroId NVARCHAR(100) = CONVERT(NVARCHAR(100), @PolizaId);
        DECLARE @CorrelacionId UNIQUEIDENTIFIER = NEWID();
        DECLARE @Cambios NVARCHAR(MAX) = (
            SELECT @Aseguradora AS Aseguradora, @NumeroPoliza AS NumeroPoliza, @Cobertura AS Cobertura,
                   @VigenciaDesde AS VigenciaDesde, @VigenciaHasta AS VigenciaHasta, JSON_QUERY(@Trabajadores) AS Trabajadores
            FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
        );
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId, @CorrelacionId, N'PolizaSctrHYT',
             @RegistroId, @Accion, @CambiosJson = @Cambios;
        COMMIT TRANSACTION;
        SELECT @PolizaId AS PolizaId;
    END TRY
    BEGIN CATCH
        IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END;
GO
