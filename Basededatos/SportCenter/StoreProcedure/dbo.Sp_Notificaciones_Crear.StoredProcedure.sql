USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 14/04/2026 | Crea notificacion de negocio para eventos operativos (reservas cliente web).
CREATE OR ALTER PROCEDURE [dbo].[Sp_Notificaciones_Crear]
    @NegocioId INT,
    @Tipo NVARCHAR(40),
    @Titulo NVARCHAR(120),
    @Mensaje NVARCHAR(300),
    @Entidad NVARCHAR(40) = NULL,
    @EntidadId INT = NULL,
    @UrlDestino NVARCHAR(300) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT INTO dbo.NegocioNotificaciones
        (
            NegocioId, Tipo, Titulo, Mensaje, Entidad, EntidadId, UrlDestino,
            Leida, FechaRegistroUtc
        )
        VALUES
        (
            @NegocioId,
            COALESCE(NULLIF(LTRIM(RTRIM(@Tipo)), N''), N'GENERAL'),
            LEFT(COALESCE(NULLIF(LTRIM(RTRIM(@Titulo)), N''), N'Notificacion'), 120),
            LEFT(COALESCE(NULLIF(LTRIM(RTRIM(@Mensaje)), N''), N'Sin detalle'), 300),
            NULLIF(LTRIM(RTRIM(@Entidad)), N''),
            @EntidadId,
            NULLIF(LTRIM(RTRIM(@UrlDestino)), N''),
            0,
            SYSUTCDATETIME()
        );

        SELECT CAST(SCOPE_IDENTITY() AS INT) AS Id;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
