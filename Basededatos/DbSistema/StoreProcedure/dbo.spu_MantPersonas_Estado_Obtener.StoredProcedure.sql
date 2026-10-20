-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Obtiene los estados comerciales, relacionados y personales de una entidad.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantPersonas_Estado_Obtener
    @IdEmpresa CHAR(2),
    @IdPersona CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            CASE WHEN C.IdCliente IS NULL THEN 0 ELSE 1 END AS EsCliente,
            CASE WHEN PR.IdProveedor IS NULL THEN 0 ELSE 1 END AS EsProveedor,
            CASE WHEN ER.IdPersona IS NULL THEN 0 ELSE 1 END AS EsRelacionada,
            ISNULL(ER.IndMatriz, '0') AS IndMatriz,
            ISNULL(ER.IndSubSidiaria, '0') AS IndSubSidiaria,
            ISNULL(ER.IndAsociada, '0') AS IndAsociada,
            CASE WHEN PE.IdPersona IS NULL THEN 0 ELSE 1 END AS EsPersonal,
            ISNULL(PE.IndPersonal, '0') AS IndPersonal,
            ISNULL(PE.IndSocio, '0') AS IndSocio
        FROM (SELECT @IdEmpresa AS IdEmpresa, @IdPersona AS IdPersona) AS I
        LEFT JOIN dbo.Clientes AS C
            ON I.IdEmpresa = C.IdEmpresa
           AND I.IdPersona = C.IdCliente
        LEFT JOIN dbo.Proveedores AS PR
            ON I.IdEmpresa = PR.IdEmpresa
           AND I.IdPersona = PR.IdProveedor
        LEFT JOIN dbo.EmpresasRelacionadas AS ER
            ON I.IdEmpresa = ER.IdEmpresa
           AND I.IdPersona = ER.IdPersona
        LEFT JOIN dbo.Personal AS PE
            ON I.IdEmpresa = PE.IdEmpresa
           AND I.IdPersona = PE.IdPersona;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000);
        DECLARE @ErrorSeverity INT;
        DECLARE @ErrorState INT;

        SELECT
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
