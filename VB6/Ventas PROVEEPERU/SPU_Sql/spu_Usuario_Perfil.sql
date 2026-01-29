
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Usuario_Perfil]
@idpersona		CHAR(8),
@idEmpresa		CHAR(2),
@idSucursal		CHAR(8),
@sistema		VARCHAR(10)
AS
BEGIN

select 

A.GlsPerfil,
C.GlsPersona,
d.GlsEmpresa,
(select GlsPersona from personas where idpersona = @idSucursal) AS glsSucursal

from Perfil A 
Inner Join PerfilesPorUsuario B 
On A.IdEmpresa = B.IdEmpresa 
And A.IdPerfil = B.IdPerfil 
And A.CodSistema = B.CodSistema
inner join personas c
on b.idUsuario = c.idPersona
INNER JOIN empresas d
on d.idEmpresa = a.idEmpresa
where B.IdUsuario = @idpersona 
and A.IdEmpresa = @idEmpresa 
And B.CodSistema = @sistema

END;