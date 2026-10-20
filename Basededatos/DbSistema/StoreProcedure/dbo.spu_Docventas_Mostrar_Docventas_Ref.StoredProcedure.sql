/****** Object:  StoredProcedure [dbo].[spu_Docventas_Mostrar_Docventas_Ref]    Script Date: 13/10/2026 21:34:11 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:        FRANCO LARA / Codex
-- Create date:   17/10/2026
-- Description:   Carga las referencias del documento de forma idempotente y sin fijar la base de despliegue.
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[spu_Docventas_Mostrar_Docventas_Ref]
@idEmpresa     CHAR(2),
@idSucursal    VARCHAR(8),
@idDocumento   CHAR(2),
@idSerie	   CHAR(4),
@iddocventas   CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY

	Select R.Item,R.TipoDocReferencia IdDocumento,D.GlsDocumento,R.NumDocReferencia IdNumDoc,R.SerieDocReferencia IdSerie,R.IndImportado 
	From DocReferencia R 
	Inner Join Documentos D 
	On R.TipoDocReferencia = D.IdDocumento 
	Where R.IdEmpresa = @idEmpresa 
	And R.IdSucursal = @idSucursal 
	And R.TipoDocOrigen = @idDocumento
	And R.SerieDocOrigen = @idSerie
	And R.NumDocOrigen = @iddocventas 
	Order By Item
	;

    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000)
        DECLARE @ErrorSeverity INT
        DECLARE @ErrorState INT

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE()

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState)
    END CATCH

END
GO
