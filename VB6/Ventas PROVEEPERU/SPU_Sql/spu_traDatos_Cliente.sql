
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_traDatos_Cliente]
@idEmpresa   CHAR(2),
@idCliente   CHAR(8)
AS
BEGIN

DECLARE @idserie CHAR(4)

SELECT p.ruc,(p.direccion + ' ' + IsNull(u.glsUbigeo,'') + ' ' + IsNull(d.glsUbigeo,'')) direccion,p.GlsPersona,p.direccionEntrega ,
c.idVendedorCampo,
c.idEmpTrans,
c.Val_Dscto,
f.idFormaPago
FROM personas p 
INNER JOIN Clientes c
on p.idPersona =c.idCliente
and c.idEmpresa = @idEmpresa
left join clientesformapagos f
on f.idcliente = c.idCliente
and c.idEmpresa = f.idempresa
LEFT JOIN ubigeo u 
ON P.IdPais = U.IdPais And P.idDistrito = u.idDistrito 
LEFT JOIN ubigeo d 
ON u.IdPais = d.IdPais And 
left(u.idDistrito,2) = d.idDpto 
AND d.idProv = '00' 
AND d.idDist = '00' 
Where p.idPersona = @idCliente

;

END;
