
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Recupera_Direccion_Cliente]
@Tipo			TinyInt,
@VarEmpresa		CHAR(2),
@VarPersona		CHAR(8),
@VarIdTieNda	VARCHAR(20)
AS
BEGIN

IF @Tipo = 1 BEGIN

	SELECT p.ruc, (p.direccion + ' ' + IsNull(u.glsUbigeo,'') + ' ' + IsNull(prov.glsubigeo,'') + ' ' + IsNull(d.glsUbigeo,'')) as direccion,p.GlsPersona,p.direccionEntrega 
	FROM personas p 
	LEFT JOIN ubigeo u 
	ON P.idDistrito = u.idDistrito 
	AND p.idPais = u.idPais 
	LEFT JOIN ubigeo d 
	ON left(u.idDistrito,2) = d.idDpto 
	AND d.idProv = '00' 
	AND d.idDist = '00' 
	AND p.idPais = d.idPais 
	LEFT JOIN ubigeo prov 
	ON left(d.idDistrito,2) = prov.idDpto 
	AND SUBSTRING(u.idDistrito,3,2) = prov.idProv 
	AND prov.idDist = '00' 
	AND prov.idPais = d.idPais 
	Where p.idPersona = @VarPersona

END

IF @Tipo = 2 BEGIN

	SELECT p.ruc, (t.GlsDireccion + ' ' + IsNull(u.glsUbigeo,'') + ' ' + IsNull(prov.glsubigeo,'') + ' ' + IsNull(d.glsUbigeo,'')) as direccion,p.GlsPersona--,p.direccionEntrega 
	FROM tiendascliente T
	INNER JOIN personas p
	ON T.idPersona = p.idPersona
	LEFT JOIN ubigeo u 
	ON t.idDistrito = u.idDistrito 
	AND t.idPais = u.idPais 
	LEFT JOIN ubigeo d 
	ON left(u.idDistrito,2) = d.idDpto 
	AND d.idProv = '00' 
	AND d.idDist = '00' 
	AND t.idPais = d.idPais 
	LEFT JOIN ubigeo prov 
	ON left(d.idDistrito,2) = prov.idDpto 
	AND SUBSTRING(u.idDistrito,3,2) = prov.idProv 
	AND prov.idDist = '00' 
	AND prov.idPais = d.idPais 
	Where t.idPersona = @VarPersona
	and t.idtdacli = @VarIdTieNda
	AND T.idEmpresa = @VarEmpresa
END


END;
