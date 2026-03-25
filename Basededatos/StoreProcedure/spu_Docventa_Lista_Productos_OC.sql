
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE  OR ALTER PROCEDURE [dbo].[spu_Docventa_Lista_Productos_OC] -- EXEC spu_Docventa_Lista_Productos_OC '01',''
@idEmpresa		  CHAR(2),
@Busqueda		  VARCHAR(300)
AS
BEGIN

DECLARE @Fragmentos TABLE (Fragmento NVARCHAR(100));

INSERT INTO @Fragmentos (Fragmento)
SELECT value
FROM STRING_SPLIT(@Busqueda, '%')
WHERE value <> '';

SELECT IdProducto cod,GlsProducto des ,idFabricante,GlsUm 
From Productos p 
Inner Join  UnidadMedida u   On p.idUMCompra = u.idUm 
WHERE idEmpresa = @idEmpresa 
AND estProducto = 'A' 
AND (
	P.IdProducto like @Busqueda OR idFabricante like @Busqueda OR GlsUm like @Busqueda
	OR
	NOT EXISTS (
	SELECT 1
	FROM @Fragmentos f
	WHERE p.GlsProducto NOT LIKE '%' + f.Fragmento + '%'
				)
	)
order by 2
;

END;
GO
