/*
Firma: Codex - 07/04/2026
Descripcion: Ajusta Sp_Reservas_Cotizar para devolver politica de confirmacion incluso sin tarifa configurada y permitir precio manual en modal.
*/

USE [DbSportCenter]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Firma: Codex - 07/04/2026 | Cotiza precio de reserva segun tarifa horaria, promociones activas y politica de pago del negocio; si no hay tarifa devuelve politica y permite precio manual.
CREATE OR ALTER PROCEDURE dbo.Sp_Reservas_Cotizar
    @NegocioId INT,
    @EspacioDeportivoId INT,
    @Fecha DATE,
    @HoraInicio TIME,
    @HoraFin TIME
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        IF @HoraFin <= @HoraInicio
            RAISERROR('La hora fin debe ser mayor que la hora inicio.', 16, 1);

        DECLARE @DuracionMinutos INT = DATEDIFF(MINUTE, @HoraInicio, @HoraFin);
        IF @DuracionMinutos NOT IN (30, 60)
            RAISERROR('Solo se permite reservas de 30 o 60 minutos.', 16, 1);

        DECLARE @SedeId INT;
        DECLARE @DiaSemana INT = (DATEDIFF(DAY, '19000101', @Fecha) % 7) + 1;

        SELECT @SedeId = s.Id
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        WHERE e.Id = @EspacioDeportivoId
          AND s.NegocioId = @NegocioId
          AND e.Estado = 1
          AND s.Activo = 1;

        IF @SedeId IS NULL
            RAISERROR('El espacio deportivo no esta disponible para este negocio.', 16, 1);

        DECLARE @PrecioHora DECIMAL(10,2);
        SELECT TOP 1
            @PrecioHora = t.Precio
        FROM dbo.Tarifas t
        WHERE t.EspacioDeportivoId = @EspacioDeportivoId
          AND t.Activa = 1
          AND t.DiaSemana = @DiaSemana
          AND @HoraInicio >= t.HoraInicio
          AND @HoraFin <= t.HoraFin
        ORDER BY t.HoraInicio DESC;

        DECLARE @TieneTarifa BIT = CASE WHEN @PrecioHora IS NULL THEN 0 ELSE 1 END;
        DECLARE @PrecioBase DECIMAL(10,2) = CASE WHEN @TieneTarifa = 1 THEN ROUND(@PrecioHora * (@DuracionMinutos / 60.0), 2) ELSE 0 END;
        DECLARE @DescuentoPct DECIMAL(5,2) = 0;

        IF @TieneTarifa = 1
        BEGIN
            SELECT TOP 1
                @DescuentoPct = p.PorcentajeDescuento
            FROM dbo.PromocionesHorario p
            WHERE p.NegocioId = @NegocioId
              AND p.Activo = 1
              AND @Fecha BETWEEN p.FechaInicio AND p.FechaFin
              AND @HoraInicio >= p.HoraInicio
              AND @HoraFin <= p.HoraFin
              AND (p.EspacioDeportivoId IS NULL OR p.EspacioDeportivoId = @EspacioDeportivoId)
              AND (p.SedeId IS NULL OR p.SedeId = @SedeId)
            ORDER BY
                CASE
                    WHEN p.EspacioDeportivoId = @EspacioDeportivoId THEN 3
                    WHEN p.SedeId = @SedeId THEN 2
                    ELSE 1
                END DESC,
                p.PorcentajeDescuento DESC,
                p.Id DESC;
        END

        DECLARE @PrecioFinal DECIMAL(10,2) = ROUND(@PrecioBase * (1 - (COALESCE(@DescuentoPct, 0) / 100.0)), 2);

        DECLARE @MonedaNombre NVARCHAR(80) = N'PEN';
        DECLARE @MonedaSimbolo NVARCHAR(10) = N'S/';
        DECLARE @PoliticaConfirmacionPago TINYINT = 0;
        DECLARE @PorcentajeAdelantoMinimo DECIMAL(5,2) = NULL;

        SELECT
            @PoliticaConfirmacionPago = COALESCE(n.PoliticaConfirmacionPago, 0),
            @PorcentajeAdelantoMinimo = n.PorcentajeAdelantoMinimo,
            @MonedaNombre = COALESCE(m.Nombre, N'PEN'),
            @MonedaSimbolo = COALESCE(NULLIF(LTRIM(RTRIM(m.Simbolo)), N''), N'S/')
        FROM dbo.Negocios n
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        WHERE n.Id = @NegocioId;

        SELECT
            CAST(CASE WHEN @TieneTarifa = 1 THEN N'Tarifa calculada correctamente.' ELSE N'No existe tarifa configurada para el horario seleccionado. Puedes ingresar el precio manualmente.' END AS NVARCHAR(200)) AS Mensaje,
            @PrecioBase AS PrecioBase,
            COALESCE(@DescuentoPct, 0) AS DescuentoPct,
            @PrecioFinal AS PrecioFinal,
            @MonedaSimbolo AS MonedaSimbolo,
            @MonedaNombre AS MonedaNombre,
            CAST(COALESCE(@PoliticaConfirmacionPago, 0) AS TINYINT) AS PoliticaConfirmacionPago,
            @PorcentajeAdelantoMinimo AS PorcentajeAdelantoMinimo;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
GO