-- =============================================
-- Author:        FRANCO LARA
-- Create date:   03/11/2026
-- Description:   Prepara la base independiente de HYT si no existe,
--                sin modificar bases existentes ni crear credenciales.
-- =============================================
USE [master];
GO
SET NOCOUNT ON;

-- CREATE DATABASE debe ejecutarse fuera de una transaccion explicita.
IF DB_ID(N'SistemaWebHYT') IS NULL
BEGIN
    EXEC(N'CREATE DATABASE [SistemaWebHYT];');
    PRINT N'Base SistemaWebHYT creada.';
END
ELSE
BEGIN
    PRINT N'La base SistemaWebHYT ya existe. No se realizaron cambios.';
END;
GO
