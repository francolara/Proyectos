-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/04/2026
-- Description:   Agrega campo ConsideracionesReserva en dbo.Sedes.
-- =============================================

IF COL_LENGTH('dbo.Sedes', 'ConsideracionesReserva') IS NULL
BEGIN
    ALTER TABLE dbo.Sedes
        ADD ConsideracionesReserva NVARCHAR(2000) NULL;
END;