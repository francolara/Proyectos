-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista clientes para el mantenimiento.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantClientes_Listar
    @IdEmpresa CHAR(2),
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        SELECT
            C.IdCliente,
            P.GlsPersona,
            CASE WHEN P.TipoPersona = '01001' THEN 'Natural' ELSE 'Juridica' END AS TipoPersona,
            P.Ruc,
            P.Direccion + ', ' + ISNULL(U.GlsUbigeo, '') AS Direccion,
            C.IdClienteInterno,
            C.Val_Dscto
        FROM dbo.Clientes AS C
        INNER JOIN dbo.Personas AS P
            ON C.IdCliente = P.IdPersona
        LEFT JOIN dbo.Ubigeo AS U
            ON P.IdDistrito = U.IdDistrito
           AND P.IdPais = U.IdPais
        WHERE C.IdEmpresa = @IdEmpresa
          AND (
                @Busqueda = ''
                OR P.GlsPersona LIKE '%' + @Busqueda + '%'
                OR C.IdCliente LIKE '%' + @Busqueda + '%'
                OR P.Ruc LIKE '%' + @Busqueda + '%'
              )
        ORDER BY C.IdCliente;
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
