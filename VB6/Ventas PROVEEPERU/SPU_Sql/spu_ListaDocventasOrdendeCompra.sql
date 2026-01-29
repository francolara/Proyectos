
CREATE OR ALTER PROCEDURE spu_ListaDocventasOrdendeCompra
    @varEmpresa CHAR(2),
    @varSucursal CHAR(8),
    @varidDocumento CHAR(2),
    @varAnnioMov INT,
    @varMesMov INT,
    @varParamBusqueda VARCHAR(500)
AS
BEGIN
    SET NOCOUNT ON;

    IF @varParamBusqueda = '' 
    BEGIN
        SET @varParamBusqueda = '%%';
    END

    IF OBJECT_ID('tempdb..#TmpReferenciasOC') IS NOT NULL DROP TABLE #TmpReferenciasOC;

    CREATE TABLE #TmpReferenciasOC (
        IdEmpresa CHAR(2) NOT NULL DEFAULT '',
        IdSucursal VARCHAR(8) NOT NULL DEFAULT '',
        TipoDocOrigen VARCHAR(2) NOT NULL DEFAULT '',
        SerieDocOrigen VARCHAR(4) NOT NULL DEFAULT '',
        NumDocOrigen VARCHAR(8) NOT NULL DEFAULT '',
        Referencia VARCHAR(250) NOT NULL DEFAULT '',
        PRIMARY KEY(IdSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen)
    );

    INSERT INTO #TmpReferenciasOC
    SELECT IdEmpresa, IdSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen,
        Referencia AS Referencia
    FROM (
        SELECT * 
        FROM (
            SELECT A.IdEmpresa, A.IdSucursal, A.TipoDocOrigen, A.SerieDocOrigen, A.NumDocOrigen,
            
			STRING_AGG(B.AbreDocumento + A.SerieDocReferencia + A.NumDocReferencia, ' - ') AS Referencia
            FROM DocReferencia A
            INNER JOIN Documentos B ON A.TipoDocReferencia = B.IdDocumento
            WHERE A.IdEmpresa = @varEmpresa AND A.TipoDocOrigen = @varidDocumento
			GROUP BY A.IdEmpresa, A.IdSucursal, A.TipoDocOrigen, A.SerieDocOrigen, A.NumDocOrigen

        ) AS A
    ) AS A
   -- GROUP BY IdSucursal, TipoDocOrigen, SerieDocOrigen, NumDocOrigen;

    IF OBJECT_ID('tempdb..#TmpReferenciasOCVales') IS NOT NULL DROP TABLE #TmpReferenciasOCVales;

    CREATE TABLE #TmpReferenciasOCVales (
        IdEmpresa CHAR(2) NOT NULL DEFAULT '',
        IdSucursal VARCHAR(8) NOT NULL DEFAULT '',
        TipoDocReferencia VARCHAR(2) NOT NULL DEFAULT '',
        SerieDocReferencia VARCHAR(4) NOT NULL DEFAULT '',
        NumDocReferencia VARCHAR(8) NOT NULL DEFAULT '',
        ReferenciaVale VARCHAR(500) NOT NULL DEFAULT '',
        PRIMARY KEY(IdSucursal, TipoDocReferencia, SerieDocReferencia, NumDocReferencia)
    );

    INSERT INTO #TmpReferenciasOCVales
    SELECT IdEmpresa, IdSucursal, TipoDocReferencia, SerieDocReferencia, NumDocReferencia,
       Referencia AS Referencia
    FROM (
        SELECT * 
        FROM (
            SELECT A.IdEmpresa, A.IdSucursal, A.TipoDocReferencia, A.SerieDocReferencia, A.NumDocReferencia,
               
			   STRING_AGG(A.NumDocOrigen, ' - ') AS Referencia

            FROM DocReferencia A
            INNER JOIN Documentos B ON A.TipoDocOrigen = B.IdDocumento
            WHERE A.IdEmpresa = @varEmpresa AND A.TipoDocReferencia = @varidDocumento AND A.TipoDocOrigen = '88'
			GROUP BY A.IdEmpresa, A.IdSucursal, A.TipoDocReferencia, A.SerieDocReferencia, A.NumDocReferencia
        ) AS A
    ) AS A
    --GROUP BY IdSucursal, TipoDocReferencia, SerieDocReferencia, NumDocReferencia;

    SELECT CONCAT(idDocumento, idDocVentas, idSerie) AS Item, idDocVentas, idSerie, idPerCliente, GlsCliente,
        RUCCliente, CONVERT(VARCHAR,FecEmision, 103) AS FecEmision, estDocVentas,
        TotalValorVenta AS TotalPrecioVenta, ISNULL(C.Referencia, '') AS Referencia, A.IdCentroCosto,
        CASE WHEN A.idMoneda = 'PEN' THEN 'Soles' ELSE 'Dolares' END AS idMoneda, nombres AS GlsUser, 
        RIGHT('00' + CAST(MONTH(A.FecEmision) AS VARCHAR(2)) ,2) AS Mes, A.GlsPlaca, ISNULL(A.IndCerrado, '0') AS IndCerrado, ISNULL(D.ReferenciaVale, '') AS ReferenciaVale
    FROM docventas A
    LEFT JOIN #TmpReferenciasOC C ON A.IdEmpresa = C.IdEmpresa AND A.IdSucursal = C.IdSucursal AND A.IdDocumento = C.TipoDocOrigen
        AND A.IdSerie = C.SerieDocOrigen AND A.IdDocVentas = C.NumDocOrigen
    LEFT JOIN #TmpReferenciasOCVales D ON A.IdEmpresa = D.IdEmpresa AND A.IdSucursal = D.IdSucursal AND A.IdDocumento = D.TipoDocReferencia
        AND A.IdSerie = D.SerieDocReferencia AND A.IdDocVentas = D.NumDocReferencia
    LEFT JOIN Personas p ON A.idPerVendedor = p.idPersona
    WHERE A.idEmpresa = @varEmpresa
    AND A.idSucursal = @varSucursal
    AND A.idDocumento = @varidDocumento
    AND YEAR(A.FecEmision) = @varAnnioMov
    AND MONTH(FecEmision) LIKE @varMesMov
   -- AND (@VarUsuario = '' OR A.IdUsuarioReg = @VarUsuario)
    AND (GlsCliente LIKE @varParamBusqueda OR idDocVentas LIKE @varParamBusqueda OR RUCCliente LIKE @varParamBusqueda
    OR idPerCliente LIKE @varParamBusqueda OR nombres LIKE @varParamBusqueda)
    ORDER BY MONTH(FecEmision), idSerie, idDocVentas;

    SET NOCOUNT OFF;
END
GO
