
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Parametros]
@varEmpresa CHAR(2)
AS
BEGIN

Select GlsParametro,ValParametro From parametros Where idEmpresa = @varEmpresa;

END;
