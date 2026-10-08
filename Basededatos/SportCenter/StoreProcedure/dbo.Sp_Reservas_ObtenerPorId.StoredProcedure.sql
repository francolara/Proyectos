
GO
/****** Object:  StoredProcedure [dbo].[Sp_Reservas_ObtenerPorId]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 04_Reservas_Pagos_Comprobantes.sql (linea 104)
-- Firma: Codex - 07/04/2026 | Incluye Comentario en consulta de detalle de reserva para pop-up.
-- Firma: FRANCO LARA - 17/09/2026 | Conserva CodigoCupon en el orden esperado y expone trazabilidad de usuario y fecha de registro/actualizacion para el pop-up de reservas.
-- Firma: Codex - 01/10/2026 | Expone NumeroPorNegocio para identificar la reserva en el titulo visible del pop-up sin mostrar el Id tecnico.
-- Firma: FRANCO LARA - 06/10/2026 | Expone CodigoMoneda y simbolo historicos de la reserva para su edicion.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Reservas_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            r.Id,
            r.EspacioDeportivoId,
            r.ClienteId,
            r.Fecha,
            r.HoraInicio,
            r.HoraFin,
            r.Total,
            r.Adelanto,
            r.Estado,
            r.Comentario,
            r.CodigoCuponAplicado AS CodigoCupon,
            r.UsuarioCreacion,
            CAST(r.FechaRegistro AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaRegistro,
            r.UsuarioActualizacion,
            CAST(r.FechaActualizacion AT TIME ZONE 'UTC' AT TIME ZONE 'SA Pacific Standard Time' AS DATETIME2) AS FechaActualizacion,
            r.NumeroPorNegocio,
            r.CodigoMoneda,
            COALESCE(ms.Simbolo, r.CodigoMoneda) AS MonedaSimbolo
        FROM dbo.Reservas r
        INNER JOIN dbo.EspaciosDeportivos e ON e.Id = r.EspacioDeportivoId
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.MonedasSuperMaestro ms ON ms.Codigo = r.CodigoMoneda
        WHERE r.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO
