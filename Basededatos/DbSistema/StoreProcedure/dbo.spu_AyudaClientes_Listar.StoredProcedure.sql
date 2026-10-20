-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 16/10/2026
-- Description: Busca clientes y recupera sus datos principales para la ayuda independiente.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_AyudaClientes_Listar
    @IdEmpresa CHAR(2),
    @Texto VARCHAR(200) = '',
    @IdUsuario CHAR(8) = '',
    @FiltrarPorVendedor BIT = 0,
    @IdCliente CHAR(8) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Texto = LTRIM(RTRIM(ISNULL(@Texto, '')));
        SET @IdCliente = LTRIM(RTRIM(ISNULL(@IdCliente, '')));

        SELECT
            P.IdPersona AS cod,
            P.GlsPersona AS des,
            ISNULL(P.Ruc, '') AS Ruc,
            ISNULL(P.Telefonos, '') AS Telefonos,
            LTRIM(RTRIM(CONCAT(
                ISNULL(P.Direccion, ''), ' ',
                ISNULL(U.GlsUbigeo, ''), ' ',
                ISNULL(D.GlsUbigeo, '')
            ))) AS Direccion,
            ISNULL(P.GlsNombreComercial, '') AS GlsNombreComercial
        FROM dbo.Personas AS P
        INNER JOIN dbo.Clientes AS C
            ON C.IdEmpresa = @IdEmpresa
           AND C.IdCliente = P.IdPersona
        LEFT JOIN dbo.Ubigeo AS U
            ON U.IdPais = P.IdPais
           AND U.IdDistrito = P.IdDistrito
        LEFT JOIN dbo.Ubigeo AS D
            ON D.IdPais = U.IdPais
           AND D.IdDpto = U.IdDpto
           AND D.IdProv = '00'
           AND D.IdDist = '00'
        WHERE (@FiltrarPorVendedor = 0 OR C.IdVendedorCampo = @IdUsuario)
          AND
          (
              (@IdCliente <> '' AND P.IdPersona = @IdCliente)
              OR
              (
                  @IdCliente = ''
                  AND @Texto <> ''
                  AND
                  (
                      P.GlsPersona LIKE '%' + @Texto + '%'
                      OR C.IdClienteInterno LIKE '%' + @Texto + '%'
                      OR P.Ruc LIKE '%' + @Texto + '%'
                      OR P.GlsNombreComercial LIKE '%' + @Texto + '%'
                  )
              )
          )
        ORDER BY P.IdPersona;
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
