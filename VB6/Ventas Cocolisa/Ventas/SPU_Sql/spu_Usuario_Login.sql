
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Usuario_Login]
@varUsuario		CHAR(20),
@idSucursal		CHAR(8)
AS
BEGIN

SELECT a.varUsuario,a.idUsuario,a.varPass ,a.idPerfil
FROM usuarios a
INNER JOIN sucursalesempresa b
on a.idUsuario = b.idUsuario
and a.idEmpresa = b.idEmpresa
WHERE  a.varUsuario = @varUsuario;


END;
