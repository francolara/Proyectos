
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Sucursales]
@varEmpresa CHAR(2)
AS
BEGIN

SELECT a.idSucursal ,b.GlsPersona 
FROM sucursales a
inner join personas  b on a.idSucursal = b.idPersona
WHERE  a.idEmpresa = @varEmpresa;

END;
