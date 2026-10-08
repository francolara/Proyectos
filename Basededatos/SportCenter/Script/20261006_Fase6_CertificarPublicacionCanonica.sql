-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Fase 6. Certifica que la publicacion usa exclusivamente contratos canonicos.
-- =============================================
-- Es una certificacion de solo lectura. Debe ejecutarse despues de publicar los SP y la aplicacion.

SET NOCOUNT ON;

BEGIN TRY
    IF OBJECT_ID(N'dbo.Sp_Sistema_ValidarContratoCanonico', N'P') IS NULL
        RAISERROR('Fase 6 incompleta: falta el validador canonico.', 16, 1);

    EXEC dbo.Sp_Sistema_ValidarContratoCanonico @ValidarDatos = 1;

    DECLARE @ObjetosCriticos TABLE
    (
        Nombre SYSNAME NOT NULL PRIMARY KEY
    );

    INSERT INTO @ObjetosCriticos (Nombre)
    VALUES
        (N'Sp_ConfiguracionClub_Actualizar'),
        (N'Sp_ConfiguracionClub_Obtener'),
        (N'Sp_Reservas_Crear'),
        (N'Sp_Reservas_Actualizar'),
        (N'Sp_Pagos_Crear'),
        (N'Sp_Comprobantes_Crear'),
        (N'Sp_Comprobantes_ObtenerPorId'),
        (N'Sp_Comprobantes_ObtenerVisualizacion'),
        (N'Sp_Reportes_IngresosPorDia'),
        (N'Sp_Reportes_ReservasPorDia'),
        (N'Sp_Reportes_OcupacionPorEspacio'),
        (N'Sp_Reportes_ResumenOperativo'),
        (N'Sp_Reportes_ResumenCobranza'),
        (N'Sp_Reportes_DetallePagos'),
        (N'Sp_Reportes_DetalleReservas');

    IF EXISTS
    (
        SELECT 1
        FROM @ObjetosCriticos oc
        LEFT JOIN sys.procedures p
            ON p.schema_id = SCHEMA_ID(N'dbo')
           AND p.name = oc.Nombre
        WHERE p.object_id IS NULL
    )
        RAISERROR('Fase 6 incompleta: falta publicar uno o mas procedimientos criticos.', 16, 1);

    IF EXISTS
    (
        SELECT 1
        FROM sys.sql_expression_dependencies sed
        INNER JOIN sys.objects o ON o.object_id = sed.referencing_id
        INNER JOIN sys.columns c
            ON c.object_id = sed.referenced_id
           AND c.column_id = sed.referenced_minor_id
        WHERE o.type IN (N'P', N'V', N'FN', N'IF', N'TF', N'TR')
          AND
          (
              (sed.referenced_id = OBJECT_ID(N'dbo.Negocios') AND c.name = N'MonedaId')
              OR
              (
                  sed.referenced_id = OBJECT_ID(N'dbo.ComprobantesElectronicos')
                  AND c.name IN (N'TipoMoneda', N'TipoComprobante')
              )
          )
    )
        RAISERROR('Fase 6 incompleta: existen modulos dependientes de columnas heredadas.', 16, 1);

    SELECT
        p.name AS Procedimiento,
        p.modify_date AS FechaModificacion
    FROM @ObjetosCriticos oc
    INNER JOIN sys.procedures p
        ON p.schema_id = SCHEMA_ID(N'dbo')
       AND p.name = oc.Nombre
    ORDER BY p.name;

    SELECT
        N'FASE_6_CERTIFICADA' AS Resultado,
        N'CANONICO_MAESTROS_V2' AS VersionContrato,
        DB_NAME() AS BaseDatos,
        SYSUTCDATETIME() AS FechaCertificacionUtc;
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
END CATCH;
