-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Lista vendedores por empresa y texto de búsqueda.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantVendedores_Listar
    @IdEmpresa CHAR(2),
    @Texto VARCHAR(100) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            V.IdVendedor,
            P.GlsPersona,
            P.Direccion + ', ' + ISNULL(U.GlsUbigeo, '') AS Direccion
        FROM dbo.Vendedores AS V
        INNER JOIN dbo.Personas AS P
            ON P.IdPersona = V.IdVendedor
        LEFT JOIN dbo.Ubigeo AS U
            ON U.IdDistrito = P.IdDistrito
           AND U.IdPais = P.IdPais
        WHERE V.IdEmpresa = @IdEmpresa
          AND (LTRIM(RTRIM(ISNULL(@Texto, ''))) = '' OR P.GlsPersona LIKE '%' + @Texto + '%')
        ORDER BY V.IdVendedor;
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
