-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/06/2026
-- Description:   Inicializa origenes contables base para todas las empresas existentes.
-- =============================================

DECLARE @IdEmpresa INT;

DECLARE empresa_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT e.IdEmpresa
FROM dbo.SEG_Empresa AS e;

OPEN empresa_cursor;

FETCH NEXT FROM empresa_cursor INTO @IdEmpresa;

WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC dbo.usp_CON_GenerarOrigenesBaseEmpresa
        @IdEmpresa = @IdEmpresa,
        @UsuarioRegistro = N'SISTEMA';

    FETCH NEXT FROM empresa_cursor INTO @IdEmpresa;
END;

CLOSE empresa_cursor;
DEALLOCATE empresa_cursor;
