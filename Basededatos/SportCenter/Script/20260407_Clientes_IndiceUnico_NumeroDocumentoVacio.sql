/*
Firma: Codex - 07/04/2026
Descripcion: Ajusta indice unico de Clientes para excluir tipo documento 0 (Doc. trib. no dom. sin RUC) del control de duplicados.
*/
USE [DbSportCenter]
GO

IF EXISTS (SELECT 1 FROM sys.indexes WHERE object_id = OBJECT_ID(N'dbo.Clientes') AND name = N'UX_Clientes_Negocio_Tipo_Numero_Activo')
BEGIN
    DROP INDEX [UX_Clientes_Negocio_Tipo_Numero_Activo] ON dbo.Clientes;
END
GO

CREATE UNIQUE NONCLUSTERED INDEX [UX_Clientes_Negocio_Tipo_Numero_Activo]
ON dbo.Clientes (NegocioId, TipoDocumento, NumeroDocumento)
WHERE Activo = 1
  AND TipoDocumento <> N'0';
GO
