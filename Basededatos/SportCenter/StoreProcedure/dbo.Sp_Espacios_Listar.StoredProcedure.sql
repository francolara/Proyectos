
GO
/****** Object:  StoredProcedure [dbo].[Sp_Espacios_Listar]    Script Date: 3/04/2026 23:18:34 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- SOURCE: 32_Usuarios_Sede_Restriccion_Filtros.sql (linea 123)
-- Firma: Codex - 13/04/2026 | Resumen de tarifas en listado de espacios agrupado por dia con rango de precios (min-max), sin detalle por franja horaria; salida incluye TieneIluminacion y Techada para badges en la UI.
-- Firma: Codex - 18/04/2026 | Incluye bandera AdministracionPrivada para identificar espacios ocultos del portal publico.
-- Firma: FRANCO LARA - 06/06/2026 | Incluye indicador y cantidad de espacios compartidos activos en el listado de espacios.
-- Firma: FRANCO LARA - 08/06/2026 | Cuenta relaciones operativas directas y de composicion sin depender solo de pares bidireccionales.
CREATE OR ALTER PROCEDURE [dbo].[Sp_Espacios_Listar]
    @NegocioId INT,
    @SedeId INT = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @SimboloMoneda NVARCHAR(10);
        SET @SimboloMoneda = N'S/';

        SELECT TOP (1) @SimboloMoneda = COALESCE(m.Simbolo, N'S/')
        FROM dbo.Negocios n
        LEFT JOIN dbo.Monedas m ON m.Id = n.MonedaId
        WHERE n.Id = @NegocioId;

        SELECT
            e.Id,
            e.Codigo,
            e.Nombre,
            s.Nombre AS Sede,
            td.Nombre AS TipoDeporte,
            ts.Nombre AS TipoSuelo,
            e.TieneIluminacion,
            e.Techada,
            CASE e.Estado WHEN 1 THEN N'Activo' WHEN 2 THEN N'EnMantenimiento' ELSE N'Inactivo' END AS Estado,
            COALESCE
            (
                NULLIF
                (
                    STUFF
                    (
                        (
                            SELECT N' - '
                                + CASE dt.DiaSemana
                                    WHEN 1 THEN N'Lunes'
                                    WHEN 2 THEN N'Martes'
                                    WHEN 3 THEN N'Miercoles'
                                    WHEN 4 THEN N'Jueves'
                                    WHEN 5 THEN N'Viernes'
                                    WHEN 6 THEN N'Sabado'
                                    WHEN 0 THEN N'Domingo'
                                    ELSE N'Dia'
                                  END
                                + N' ('
                                + @SimboloMoneda + N' '
                                + REPLACE(REPLACE(CONVERT(NVARCHAR(20), CAST(dt.PrecioMin AS DECIMAL(10,2))), N'.00', N''), N',00', N'')
                                + CASE
                                    WHEN dt.PrecioMax > dt.PrecioMin
                                        THEN N' - ' + @SimboloMoneda + N' ' + REPLACE(REPLACE(CONVERT(NVARCHAR(20), CAST(dt.PrecioMax AS DECIMAL(10,2))), N'.00', N''), N',00', N'')
                                    ELSE N''
                                  END
                                + N')'
                            FROM
                            (
                                SELECT
                                    t.DiaSemana,
                                    MIN(t.Precio) AS PrecioMin,
                                    MAX(t.Precio) AS PrecioMax
                                FROM dbo.Tarifas t
                                WHERE t.EspacioDeportivoId = e.Id
                                  AND t.Activa = 1
                                GROUP BY t.DiaSemana
                            ) dt
                            ORDER BY
                                CASE dt.DiaSemana
                                    WHEN 1 THEN 1
                                    WHEN 2 THEN 2
                                    WHEN 3 THEN 3
                                    WHEN 4 THEN 4
                                    WHEN 5 THEN 5
                                    WHEN 6 THEN 6
                                    WHEN 0 THEN 7
                                    ELSE 8
                                END
                            FOR XML PATH(''), TYPE
                        ).value('.', 'NVARCHAR(MAX)'),
                        1, 3, N''
                    ),
                    N''
                ),
                N'Sin tarifa configurada'
            ) AS TarifaResumen,
            COALESCE(e.AdministracionPrivada, 0) AS AdministracionPrivada,
            CAST(
                CASE
                    WHEN EXISTS
                    (
                        SELECT 1
                        FROM dbo.EspaciosDeportivosCompartidos ec
                        WHERE ec.Activo = 1
                          AND
                          (
                              ec.EspacioDeportivoId = e.Id
                              OR (ec.TipoRelacion = N'COMPUESTO_COMPONENTE' AND ec.EspacioRelacionadoId = e.Id)
                          )
                    ) THEN 1 ELSE 0
                END
                AS BIT
            ) AS TieneEspaciosCompartidos,
            (
                SELECT COUNT(DISTINCT RelacionadoId)
                FROM
                (
                    SELECT ec.EspacioRelacionadoId AS RelacionadoId
                    FROM dbo.EspaciosDeportivosCompartidos ec
                    WHERE ec.EspacioDeportivoId = e.Id
                      AND ec.Activo = 1

                    UNION ALL

                    SELECT ec.EspacioDeportivoId AS RelacionadoId
                    FROM dbo.EspaciosDeportivosCompartidos ec
                    WHERE ec.TipoRelacion = N'COMPUESTO_COMPONENTE'
                      AND ec.EspacioRelacionadoId = e.Id
                      AND ec.Activo = 1
                ) relaciones
            ) AS TotalEspaciosCompartidos
        FROM dbo.EspaciosDeportivos e
        INNER JOIN dbo.Sedes s ON s.Id = e.SedeId
        INNER JOIN dbo.TiposDeporte td ON td.Id = e.TipoDeporteId
        INNER JOIN dbo.TiposSuelo ts ON ts.Id = e.TipoSueloId
        WHERE s.NegocioId = @NegocioId
          AND (@SedeId IS NULL OR s.Id = @SedeId)
        ORDER BY s.Nombre, e.Nombre;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END

GO


