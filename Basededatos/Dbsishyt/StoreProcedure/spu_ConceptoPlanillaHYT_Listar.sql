-- Firma: FRANCO LARA | Create date: 10/11/2026 | Consulta el catálogo de conceptos con permisos de trabajadores.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ConceptoPlanillaHYT_Listar
    @UsuarioId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,N'TRABAJADORES.CONSULTAR')=0
            THROW 51000,N'No tiene permiso para consultar conceptos de planilla.',1;

        SELECT concepto.ConceptoId,concepto.Nombre,concepto.Naturaleza,concepto.Calculo,concepto.Valor,
               concepto.Regimen,concepto.Activo,concepto.VersionFila
        FROM dbo.ConceptoPlanillaHYT AS concepto
        ORDER BY concepto.Nombre,concepto.ConceptoId;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
