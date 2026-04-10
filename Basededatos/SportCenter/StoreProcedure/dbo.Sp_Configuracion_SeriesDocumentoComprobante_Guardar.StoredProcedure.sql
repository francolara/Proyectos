USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA
-- Create date:   09/04/2026
-- Description:   Crea/actualiza serie de comprobante por negocio y documento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.Sp_Configuracion_SeriesDocumentoComprobante_Guardar
    @NegocioId INT,
    @CodigoSunat NVARCHAR(4),
    @Serie NVARCHAR(4),
    @Activo BIT = 1,
    @Usuario NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Tributario BIT;

        SET @CodigoSunat = UPPER(LTRIM(RTRIM(@CodigoSunat)));
        SET @Serie = UPPER(LTRIM(RTRIM(@Serie)));

        IF @CodigoSunat IS NULL OR @CodigoSunat = N''
            RAISERROR('Selecciona un tipo de documento.', 16, 1);

        IF @Serie IS NULL OR LEN(@Serie) <> 4
            RAISERROR('La serie debe tener exactamente 4 caracteres.', 16, 1);

        SELECT @Tributario = t.Tributario
        FROM dbo.TiposDocumentoComprobanteSuperMaestro t
        WHERE t.CodigoSunat = @CodigoSunat
          AND t.Activo = 1
          AND t.Habilitado = 1;

        IF @Tributario IS NULL
            RAISERROR('El tipo de documento no esta habilitado.', 16, 1);

        IF @CodigoSunat = N'01' AND LEFT(@Serie, 1) <> N'F'
            RAISERROR('La serie de Factura debe iniciar con F.', 16, 1);

        IF @CodigoSunat = N'03' AND LEFT(@Serie, 1) <> N'B'
            RAISERROR('La serie de Boleta debe iniciar con B.', 16, 1);

        IF @Tributario = 0 AND LEFT(@Serie, 1) <> N'R'
            RAISERROR('La serie de documento no tributario debe iniciar con R.', 16, 1);

        IF NOT EXISTS (
            SELECT 1
            FROM dbo.NegociosTiposDocumentoComprobante ntd
            WHERE ntd.NegocioId = @NegocioId
              AND ntd.CodigoSunat = @CodigoSunat
              AND ntd.Activo = 1
        )
            RAISERROR('El documento no esta habilitado en Maestros para este negocio.', 16, 1);

        MERGE dbo.NegociosSeriesDocumentoComprobante AS tgt
        USING (SELECT @NegocioId AS NegocioId, @CodigoSunat AS CodigoSunat, @Serie AS Serie) AS src
        ON tgt.NegocioId = src.NegocioId
           AND tgt.CodigoSunat = src.CodigoSunat
           AND tgt.Serie = src.Serie
        WHEN MATCHED THEN
            UPDATE SET
                tgt.Activo = @Activo,
                tgt.FechaActualizacion = SYSUTCDATETIME(),
                tgt.UsuarioActualizacion = @Usuario
        WHEN NOT MATCHED BY TARGET THEN
            INSERT (NegocioId, CodigoSunat, Serie, Activo, FechaCreacion, UsuarioCreacion)
            VALUES (src.NegocioId, src.CodigoSunat, src.Serie, @Activo, SYSUTCDATETIME(), @Usuario);
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
