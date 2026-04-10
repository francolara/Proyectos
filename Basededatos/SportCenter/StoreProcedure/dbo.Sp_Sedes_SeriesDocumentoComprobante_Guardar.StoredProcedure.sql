USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Guarda serie de documento por sede (opcional por documento).
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Sedes_SeriesDocumentoComprobante_Guardar
    @NegocioId INT,
    @SedeId INT,
    @CodigoSunat NVARCHAR(4),
    @NegocioSerieId INT = NULL,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @CodigoSunat = UPPER(LTRIM(RTRIM(@CodigoSunat)));
        IF @CodigoSunat IS NULL OR @CodigoSunat = N''
            RAISERROR('Selecciona documento para la sede.', 16, 1);

        IF NOT EXISTS (SELECT 1 FROM dbo.Sedes WHERE Id = @SedeId AND NegocioId = @NegocioId)
            RAISERROR('La sede no pertenece al negocio.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.NegociosTiposDocumentoComprobante ntd
            INNER JOIN dbo.TiposDocumentoComprobanteSuperMaestro t ON t.CodigoSunat = ntd.CodigoSunat
            WHERE ntd.NegocioId = @NegocioId
              AND ntd.CodigoSunat = @CodigoSunat
              AND ntd.Activo = 1
              AND t.Habilitado = 1
              AND t.Activo = 1
        )
            RAISERROR('El documento no esta habilitado para el negocio.', 16, 1);

        IF @NegocioSerieId IS NULL
        BEGIN
            UPDATE dbo.SedesSeriesDocumentoComprobante
            SET Activo = 0,
                FechaActualizacion = SYSUTCDATETIME(),
                UsuarioActualizacion = @Usuario
            WHERE SedeId = @SedeId
              AND CodigoSunat = @CodigoSunat
              AND Activo = 1;
            RETURN;
        END

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.NegociosSeriesDocumentoComprobante ns
            WHERE ns.Id = @NegocioSerieId
              AND ns.NegocioId = @NegocioId
              AND ns.CodigoSunat = @CodigoSunat
              AND ns.Activo = 1
        )
            RAISERROR('La serie seleccionada no corresponde al negocio/documento.', 16, 1);

        MERGE dbo.SedesSeriesDocumentoComprobante AS tgt
        USING (SELECT @SedeId AS SedeId, @CodigoSunat AS CodigoSunat) AS src
        ON tgt.SedeId = src.SedeId AND tgt.CodigoSunat = src.CodigoSunat
        WHEN MATCHED THEN
            UPDATE SET
                tgt.NegocioSerieId = @NegocioSerieId,
                tgt.Activo = 1,
                tgt.FechaActualizacion = SYSUTCDATETIME(),
                tgt.UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED BY TARGET THEN
            INSERT (SedeId, CodigoSunat, NegocioSerieId, Activo, FechaCreacion, UsuarioCreacion)
            VALUES (@SedeId, @CodigoSunat, @NegocioSerieId, 1, SYSUTCDATETIME(), @Usuario);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
