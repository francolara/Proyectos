
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE OR ALTER  PROCEDURE [dbo].[spu_Documento]
@idDocumento CHAR(2)
AS
BEGIN

SELECT GlsDocumento,idDocumento,frameHeight FROM documentos
where idDocumento = @idDocumento

;

END;
