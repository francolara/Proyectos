
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Docventas_VariableIni]
@idEmpresa   CHAR(2),
@idDocumento CHAR(2),
@idUsuario   CHAR(8),
@VarFecha	 VARCHAR(10)
AS
BEGIN

DECLARE @idserie CHAR(4)
DECLARE @TcVenta DECIMAL(10,3)
DECLARE @idVendedor VARCHAR(8)

SELECT 
@idserie = idSerie
FROM seriexusuario a
where a.idEmpresa = @idEmpresa
and A.idUsuario = @idUsuario  
and A.idDocumento = @idDocumento


SELECT 
@TcVenta = TcVenta 
FROM TiposDeCambio
WHERE CAST(Fecha AS DATE) = CAST(@VarFecha AS DATE)


SELECT 
@idVendedor = idVendedor 
FROM vendedores
WHERE idVendedor = @idUsuario



SELECT ISNULL(@idserie,'') AS idserie, ISNULL(@TcVenta,0) AS TcVenta, ISNULL(@idVendedor,'') AS idVendedor

;

END;
