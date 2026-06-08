
GO
/****** Object:  StoredProcedure [dbo].[Sp_Espacios_ObtenerPorId]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 31_Espacios_Tarifas_Base.sql (linea 29)
-- Firma: Codex - 18/04/2026 | Devuelve bandera AdministracionPrivada para edicion y control de visibilidad publica.
-- Firma: FRANCO LARA - 26/05/2026 | Devuelve configuracion de horario por espacio deportivo (switch, dias y rango horario).
-- Firma: FRANCO LARA - 06/06/2026 | Devuelve configuracion y lista de espacios compartidos para bloqueo bidireccional.
-- Firma: FRANCO LARA - 08/06/2026 | Devuelve por separado bloqueo directo y componentes del espacio compuesto, e incluye fotos del espacio deportivo.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_ObtenerPorId]
    @NegocioId INT,
    @Id INT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SELECT
            e.Id,
            e.SedeId,
            e.TipoDeporteId,
            e.TipoSueloId,
            e.Codigo,
            e.Nombre,
            e.Capacidad,
            e.TieneIluminacion,
            e.Techada,
            e.Estado,
            (
                SELECT
                    t.DiaSemana,
                    CONVERT(NVARCHAR(8), t.HoraInicio, 108) AS HoraInicio,
                    CONVERT(NVARCHAR(8), t.HoraFin, 108) AS HoraFin,
                    t.Precio
                FROM dbo.Tarifas t
                WHERE t.EspacioDeportivoId = e.Id
                  AND t.Activa = 1
                ORDER BY t.DiaSemana, t.HoraInicio
                FOR JSON PATH
            ) AS TarifasJson,
            COALESCE(e.AdministracionPrivada, 0) AS AdministracionPrivada,
            (
                SELECT
                    CONVERT(NVARCHAR(8), t.HoraInicio, 108) AS HoraInicio,
                    CONVERT(NVARCHAR(8), t.HoraFin, 108) AS HoraFin,
                    t.Precio
                FROM dbo.TarifaFeriado t
                WHERE t.EspacioDeportivoId = e.Id
                  AND t.Activa = 1
                ORDER BY t.HoraInicio
                FOR JSON PATH
            ) AS TarifasFeriadoJson,
            e.FotoPrincipalUrl,
            e.FotosUrlsCsv,
            COALESCE(eha.ConfigurarHorarioPorEspacio, 0) AS ConfigurarHorarioPorEspacio,
            COALESCE(eha.AtiendeLunes, 1) AS AtiendeLunes,
            COALESCE(eha.AtiendeMartes, 1) AS AtiendeMartes,
            COALESCE(eha.AtiendeMiercoles, 1) AS AtiendeMiercoles,
            COALESCE(eha.AtiendeJueves, 1) AS AtiendeJueves,
            COALESCE(eha.AtiendeViernes, 1) AS AtiendeViernes,
            COALESCE(eha.AtiendeSabado, 1) AS AtiendeSabado,
            COALESCE(eha.AtiendeDomingo, 1) AS AtiendeDomingo,
            COALESCE(eha.HoraApertura, CAST('08:00' AS TIME)) AS HoraApertura,
            COALESCE(eha.HoraCierre, CAST('23:00' AS TIME)) AS HoraCierre,
            CAST(
                CASE
                    WHEN EXISTS
                    (
                        SELECT 1
                        FROM dbo.EspaciosDeportivosCompartidos ec
                        WHERE ec.Activo = 1
                          AND ec.EspacioDeportivoId = e.Id
                    ) THEN 1 ELSE 0
                END
                AS BIT
            ) AS TieneEspaciosCompartidos,
            COALESCE(
                STUFF
                (
                    (
                        SELECT N',' + CONVERT(NVARCHAR(20), ec.EspacioRelacionadoId)
                        FROM dbo.EspaciosDeportivosCompartidos ec
                        WHERE ec.EspacioDeportivoId = e.Id
                          AND ec.Activo = 1
                          AND ec.TipoRelacion = N'DIRECTO'
                        ORDER BY ec.EspacioRelacionadoId
                        FOR XML PATH(''), TYPE
                    ).value('.', 'NVARCHAR(MAX)'),
                    1, 1, N''
                ),
                N''
            ) AS EspaciosDirectosIdsCsv,
            COALESCE(
                STUFF
                (
                    (
                        SELECT N',' + CONVERT(NVARCHAR(20), ec.EspacioRelacionadoId)
                        FROM dbo.EspaciosDeportivosCompartidos ec
                        WHERE ec.EspacioDeportivoId = e.Id
                          AND ec.Activo = 1
                          AND ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
                        ORDER BY ec.EspacioRelacionadoId
                        FOR XML PATH(''), TYPE
                    ).value('.', 'NVARCHAR(MAX)'),
                    1, 1, N''
                ),
                N''
            ) AS EspaciosComponentesIdsCsv
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        LEFT JOIN dbo.EspacioHorarioAtencion eha ON eha.EspacioDeportivoId = e.Id
        WHERE e.Id = @Id
          AND s.NegocioId = @NegocioId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO
