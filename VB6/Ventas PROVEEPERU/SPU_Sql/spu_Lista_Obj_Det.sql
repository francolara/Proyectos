
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Lista_Obj_Det]
@idEmpresa   CHAR(2),
@idDocumento CHAR(2),
@idUsuario   CHAR(8)
AS
BEGIN

DECLARE @idserie CHAR(4)

SELECT 
@idserie = CASE WHEN COUNT(*) = 1 THEN '999' ELSE a.idSerie END 
FROM seriexusuario a
LEFT join objdocventas b
on a.idSerie = b.idSerie
and a.idDocumento = b.idDocumento
and (tipoObj = 'C' or tipoObj = 'T') and indVisible = 'V'
where a.idEmpresa = @idEmpresa
and A.idUsuario = @idUsuario  
and A.idDocumento = @idDocumento
GROUP BY a.idSerie

PRINT 'SERIE' + ISNULL(@idserie,'')
IF ISNULL(@idserie,'')  = '' BEGIN
	SET @idserie = '999'
END


SELECT GlsObj,etiqueta,numCol,ancho,Tipodato,Decimales  FROM objdocventas 
where idEmpresa = @idEmpresa
and idDocumento = @idDocumento 
AND idSerie = @idserie 
and  tipoObj = 'D'
and indVisible = 'V' 
ORDER BY NUMCOL

;

END;
