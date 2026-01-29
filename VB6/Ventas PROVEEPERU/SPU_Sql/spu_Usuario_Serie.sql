
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Usuario_Serie]
@idEmpresa CHAR(2),
@idUsuario CHAR(8),
@idDocumento CHAR(2)
AS
BEGIN

SELECT 
a.idSerie,a.idUsuario ,
SUM(1) AS Cantidad
FROM seriexusuario a
LEFT join objdocventas b
on a.idSerie = b.idSerie
and a.idDocumento = b.idDocumento
and (tipoObj = 'C' or tipoObj = 'T') and indVisible = 'V'
where a.idEmpresa = @idEmpresa
and A.idUsuario = @idUsuario  
and A.idDocumento = @idDocumento
GROUP BY a.idSerie,a.idUsuario 
;

END;
