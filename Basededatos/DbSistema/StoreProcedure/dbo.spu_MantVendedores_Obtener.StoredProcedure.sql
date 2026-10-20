-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Obtiene vendedor, persona, ubicación y descripciones auxiliares en una sola llamada.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantVendedores_Obtener
    @IdEmpresa CHAR(2),
    @IdVendedor CHAR(8)
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SELECT
            V.IdVendedor,
            V.IdLista,
            V.IndVendedor,
            V.IndResponsable,
            V.IndVisualizaClientes,
            V.Rpm,
            V.Nextel,
            V.IndJefe,
            V.IdJefe,
            V.IndTipoVenta,
            P.GlsPersona,
            ISNULL(P.Direccion, '') AS Direccion,
            ISNULL(P.IdDistrito, '') AS IdDistrito,
            ISNULL(U.IdDpto, '') AS IdDpto,
            ISNULL(U.IdProv, '') AS IdProv,
            ISNULL(LP.GlsLista, '') AS GlsLista,
            ISNULL(PJ.GlsPersona, '') AS GlsJefe,
            ISNULL(PA.GlsPais, '') AS GlsPais,
            ISNULL(UD.GlsUbigeo, '') AS GlsDepa,
            ISNULL(UP.GlsUbigeo, '') AS GlsProv,
            ISNULL(U.GlsUbigeo, '') AS GlsDistrito
        FROM dbo.Vendedores AS V
        INNER JOIN dbo.Personas AS P
            ON P.IdPersona = V.IdVendedor
        LEFT JOIN dbo.Ubigeo AS U
            ON U.IdDistrito = P.IdDistrito
        LEFT JOIN dbo.ListaPrecios AS LP
            ON LP.IdEmpresa = V.IdEmpresa
           AND LP.IdLista = V.IdLista
        LEFT JOIN dbo.Personas AS PJ
            ON PJ.IdPersona = V.IdJefe
        LEFT JOIN dbo.Ubigeo AS UD
            ON UD.IdDpto = U.IdDpto
           AND UD.IdProv = '00'
           AND UD.IdDist = '00'
        LEFT JOIN dbo.Ubigeo AS UP
            ON UP.IdDpto = U.IdDpto
           AND UP.IdProv = U.IdProv
           AND UP.IdDist = '00'
        OUTER APPLY
        (
            SELECT TOP (1) D.GlsDato AS GlsPais
            FROM dbo.Datos AS D
            WHERE D.IdDato = '02001'
            ORDER BY D.IdTipoDatos
        ) AS PA
        WHERE V.IdEmpresa = @IdEmpresa
          AND V.IdVendedor = @IdVendedor;
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
