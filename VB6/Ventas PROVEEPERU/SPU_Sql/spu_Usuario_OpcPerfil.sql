
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Usuario_OpcPerfil]
@Tipo			tinyint,
@idUsuario		CHAR(8),
@idEmpresa		CHAR(2),
@sistema		VARCHAR(10)
AS
BEGIN

if @Tipo  = 1 begin

	Select o.opmNum from opcionesperfil o 
	where o.idEmpresa = @idEmpresa and o.CodSistema = @sistema
	AND o.idPerfil = (select P.idPerfil From perfilesporusuario p WHERE p.idEmpresa = @idEmpresa  AND p.idUsuario = @idUsuario and CodSistema = @sistema)

end

if @Tipo  = 2 begin
	Select opmNum from opcionesmenu where opmEstado = 'N' and CodSistema = @sistema
end

END;