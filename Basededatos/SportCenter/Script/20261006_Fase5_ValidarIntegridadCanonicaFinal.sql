-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/10/2026
-- Description:   Fase 5. Audita la integridad final de monedas y comprobantes canonicos.
-- =============================================
-- Requisito: publicar primero dbo.Sp_Sistema_ValidarContratoCanonico.
-- Este script es de solo lectura y puede ejecutarse antes y despues del despliegue.

SET NOCOUNT ON;

BEGIN TRY
    IF OBJECT_ID(N'dbo.Sp_Sistema_ValidarContratoCanonico', N'P') IS NULL
        RAISERROR('Fase 5 incompleta: falta publicar Sp_Sistema_ValidarContratoCanonico.', 16, 1);

    EXEC dbo.Sp_Sistema_ValidarContratoCanonico @ValidarDatos = 1;

    SELECT
        N'Negocios' AS Entidad,
        n.CodigoMoneda,
        COUNT_BIG(*) AS Cantidad
    FROM dbo.Negocios n
    GROUP BY n.CodigoMoneda

    UNION ALL

    SELECT
        N'Reservas' AS Entidad,
        r.CodigoMoneda,
        COUNT_BIG(*) AS Cantidad
    FROM dbo.Reservas r
    GROUP BY r.CodigoMoneda

    UNION ALL

    SELECT
        N'Pagos' AS Entidad,
        p.CodigoMoneda,
        COUNT_BIG(*) AS Cantidad
    FROM dbo.Pagos p
    GROUP BY p.CodigoMoneda

    UNION ALL

    SELECT
        N'Comprobantes' AS Entidad,
        ce.CodigoMoneda,
        COUNT_BIG(*) AS Cantidad
    FROM dbo.ComprobantesElectronicos ce
    GROUP BY ce.CodigoMoneda
    ORDER BY Entidad, CodigoMoneda;

    SELECT
        ce.CodigoTipoComprobante,
        COUNT_BIG(*) AS Cantidad
    FROM dbo.ComprobantesElectronicos ce
    GROUP BY ce.CodigoTipoComprobante
    ORDER BY ce.CodigoTipoComprobante;
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
