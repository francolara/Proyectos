-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/10/2026
-- Description:   Atiende catalogos, seguridad y estados simples solicitados por FrmBusqueda sin SQL dinamico.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Ayuda_Catalogos
    @TipoAyuda VARCHAR(30),
    @Texto VARCHAR(200),
    @IdEmpresa CHAR(2) = '',
    @IdUsuario CHAR(8) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        DECLARE @IdTipoDatos VARCHAR(2);

        SET @TipoAyuda = UPPER(LTRIM(RTRIM(@TipoAyuda)));
        SET @Texto = LTRIM(RTRIM(ISNULL(@Texto, '')));
        SET @IdTipoDatos = CASE @TipoAyuda
            WHEN 'TIPOPERSONA' THEN '01'
            WHEN 'PAIS' THEN '02'
            WHEN 'TIPOPRODUCTO' THEN '06'
            WHEN 'MARCAVEHI' THEN '07'
            WHEN 'TIPOTICKET' THEN '08'
            WHEN 'PERIODOLETRA' THEN '09'
            WHEN 'TIPOINTERES' THEN '10'
            WHEN 'TIPOCUOTALETRA' THEN '11'
            WHEN 'TIPOCAPITALIZACION' THEN '12'
            WHEN 'TIPOCOBRO' THEN '13'
            WHEN 'TIPOEFECTIVO' THEN '14'
            WHEN 'TIPOSNIVEL' THEN '20'
            WHEN 'CONCEPTOSCOSTEO' THEN '26'
        END;

        IF @IdTipoDatos IS NULL
        BEGIN
            IF @TipoAyuda = 'MONEDA'
            BEGIN
                SELECT M.IdMoneda AS cod, M.GlsMoneda AS des
                FROM dbo.Monedas AS M
                WHERE M.EstMoneda = 'ACT'
                  AND (M.GlsMoneda LIKE '%' + @Texto + '%' OR M.IdMoneda LIKE '%' + @Texto + '%')
                ORDER BY M.IdMoneda;
                RETURN;
            END;

            IF @TipoAyuda = 'UM'
            BEGIN
                SELECT U.IdUM AS cod, U.AbreUM AS des
                FROM dbo.UnidadMedida AS U
                WHERE U.GlsUM LIKE '%' + @Texto + '%' OR U.IdUM LIKE '%' + @Texto + '%'
                ORDER BY U.IdUM;
                RETURN;
            END;

            IF @TipoAyuda = 'UMGLOSA'
            BEGIN
                SELECT U.IdUM AS cod, U.GlsUM AS des
                FROM dbo.UnidadMedida AS U
                WHERE U.GlsUM LIKE '%' + @Texto + '%' OR U.IdUM LIKE '%' + @Texto + '%'
                ORDER BY U.IdUM;
                RETURN;
            END;

            IF @TipoAyuda = 'TIPOFORMASPAGO'
            BEGIN
                SELECT T.IdTipoFormaPago AS cod, T.GlsTipoFormaPago AS des
                FROM dbo.TipoFormasPago AS T
                WHERE T.GlsTipoFormaPago LIKE '%' + @Texto + '%' OR T.IdTipoFormaPago LIKE '%' + @Texto + '%'
                ORDER BY T.IdTipoFormaPago;
                RETURN;
            END;

            IF @TipoAyuda = 'FORMASPAGO'
            BEGIN
                SELECT F.IdFormaPago AS cod, F.GlsFormaPago AS des
                FROM dbo.FormasPagos AS F
                WHERE F.IdEmpresa = @IdEmpresa
                  AND (F.GlsFormaPago LIKE '%' + @Texto + '%' OR F.IdFormaPago LIKE '%' + @Texto + '%')
                ORDER BY F.IdFormaPago;
                RETURN;
            END;

            IF @TipoAyuda = 'DOCUMENTOS'
            BEGIN
                SELECT D.IdDocumento AS cod, D.GlsDocumento AS des
                FROM dbo.Documentos AS D
                WHERE D.IndVentas = '*'
                  AND (D.GlsDocumento LIKE '%' + @Texto + '%' OR D.IdDocumento LIKE '%' + @Texto + '%')
                ORDER BY D.IdDocumento;
                RETURN;
            END;

            IF @TipoAyuda = 'DOCUMENTOSI'
            BEGIN
                SELECT D.IdDocumento AS cod, D.GlsDocumento AS des
                FROM dbo.Documentos AS D
                WHERE D.GlsDocumento LIKE '%' + @Texto + '%' OR D.IdDocumento LIKE '%' + @Texto + '%'
                ORDER BY D.IdDocumento;
                RETURN;
            END;

            IF @TipoAyuda = 'UNIDADPRODUC'
            BEGIN
                SELECT U.CodUnidProd AS cod, U.DescUnidad AS des
                FROM dbo.UnidadProduccion AS U
                WHERE U.IdEmpresa = @IdEmpresa
                  AND (U.DescUnidad LIKE '%' + @Texto + '%' OR U.CodUnidProd LIKE '%' + @Texto + '%')
                ORDER BY U.CodUnidProd;
                RETURN;
            END;

            IF @TipoAyuda = 'DIRRECOJO'
            BEGIN
                SELECT D.IdDirRecojo AS cod, D.GlsDirRecojo AS des
                FROM dbo.DirRecojos AS D
                WHERE D.IdEmpresa = @IdEmpresa
                  AND (D.GlsDirRecojo LIKE '%' + @Texto + '%' OR D.IdDirRecojo LIKE '%' + @Texto + '%')
                ORDER BY D.IdDirRecojo;
                RETURN;
            END;

            IF @TipoAyuda = 'MARCA'
            BEGIN
                SELECT M.IdMarca AS cod, M.GlsMarca AS des
                FROM dbo.Marcas AS M
                WHERE M.IdEmpresa = @IdEmpresa
                  AND (M.GlsMarca LIKE '%' + @Texto + '%' OR M.IdMarca LIKE '%' + @Texto + '%')
                ORDER BY M.IdMarca;
                RETURN;
            END;

            IF @TipoAyuda IN ('CONCEPTOSALIDA', 'CONCEPTOINGRESO')
            BEGIN
                SELECT C.IdConcepto AS cod, C.GlsConcepto AS des
                FROM dbo.Conceptos AS C
                WHERE C.TipoConcepto = CASE @TipoAyuda WHEN 'CONCEPTOSALIDA' THEN 'S' ELSE 'I' END
                  AND (C.GlsConcepto LIKE '%' + @Texto + '%' OR C.IdConcepto LIKE '%' + @Texto + '%')
                ORDER BY C.IdConcepto;
                RETURN;
            END;

            IF @TipoAyuda = 'PERFIL'
            BEGIN
                SELECT P.IdPerfil AS cod, P.GlsPerfil AS des
                FROM dbo.Perfil AS P
                WHERE P.IdEmpresa = @IdEmpresa
                  AND (P.GlsPerfil LIKE '%' + @Texto + '%' OR P.IdPerfil LIKE '%' + @Texto + '%')
                ORDER BY P.IdPerfil;
                RETURN;
            END;

            IF @TipoAyuda = 'EMPRESA'
            BEGIN
                SELECT E.IdEmpresa AS cod, E.GlsEmpresa AS des
                FROM dbo.Empresas AS E
                WHERE E.GlsEmpresa LIKE '%' + @Texto + '%' OR E.IdEmpresa LIKE '%' + @Texto + '%'
                ORDER BY E.IdEmpresa;
                RETURN;
            END;

            IF @TipoAyuda = 'TIPONIVEL'
            BEGIN
                SELECT T.IdTipoNivel AS cod, T.GlsTipoNivel AS des
                FROM dbo.TiposNiveles AS T
                WHERE T.IdEmpresa = @IdEmpresa
                  AND (T.GlsTipoNivel LIKE '%' + @Texto + '%' OR T.IdTipoNivel LIKE '%' + @Texto + '%')
                ORDER BY T.IdTipoNivel;
                RETURN;
            END;

            IF @TipoAyuda = 'CAJAS'
            BEGIN
                SELECT C.IdCaja AS cod, C.GlsCaja AS des
                FROM dbo.Cajas AS C
                WHERE C.IdEmpresa = @IdEmpresa
                  AND (C.GlsCaja LIKE '%' + @Texto + '%' OR C.IdCaja LIKE '%' + @Texto + '%')
                ORDER BY C.IdCaja;
                RETURN;
            END;

            IF @TipoAyuda = 'PERMISOS'
            BEGIN
                SELECT P.IdPermiso AS cod, P.GlsPermiso AS des
                FROM dbo.Permisos AS P
                WHERE P.GlsPermiso LIKE '%' + @Texto + '%' OR P.IdPermiso LIKE '%' + @Texto + '%'
                ORDER BY P.IdPermiso;
                RETURN;
            END;

            IF @TipoAyuda = 'CTACORRIENTE'
            BEGIN
                SELECT C.IdCtaCorriente AS cod, C.GlsCtaCorriente AS des
                FROM dbo.CtasCorrientes AS C
                WHERE C.IdEmpresa = @IdEmpresa
                  AND (C.GlsCtaCorriente LIKE '%' + @Texto + '%' OR C.IdCtaCorriente LIKE '%' + @Texto + '%')
                ORDER BY C.IdCtaCorriente;
                RETURN;
            END;

            IF @TipoAyuda = 'BANCO'
            BEGIN
                SELECT B.IdBanco AS cod, B.GlsBanco AS des
                FROM dbo.Bancos AS B
                WHERE B.GlsBanco LIKE '%' + @Texto + '%' OR B.IdBanco LIKE '%' + @Texto + '%'
                ORDER BY B.IdBanco;
                RETURN;
            END;

            IF @TipoAyuda = 'GRUPOSPRODUCTO'
            BEGIN
                SELECT G.IdGrupo AS cod, G.GlsGrupo AS des
                FROM dbo.GruposProducto AS G
                WHERE G.IdEmpresa = @IdEmpresa
                  AND (G.GlsGrupo LIKE '%' + @Texto + '%' OR G.IdGrupo LIKE '%' + @Texto + '%')
                ORDER BY G.IdGrupo;
                RETURN;
            END;

            IF @TipoAyuda IN ('ZONA', 'ZONAS')
            BEGIN
                SELECT Z.IdZona AS cod, Z.GlsZona AS des
                FROM dbo.Zonas AS Z
                WHERE Z.GlsZona LIKE '%' + @Texto + '%' OR Z.IdZona LIKE '%' + @Texto + '%'
                ORDER BY Z.IdZona;
                RETURN;
            END;

            IF @TipoAyuda = 'UBICACIONES'
            BEGIN
                SELECT U.IdUbicacion AS cod, U.GlsUbicacion AS des
                FROM dbo.AlmacenesUbicacion AS U
                WHERE U.IdEmpresa = @IdEmpresa
                  AND (U.GlsUbicacion LIKE '%' + @Texto + '%' OR U.IdUbicacion LIKE '%' + @Texto + '%')
                ORDER BY U.IdUbicacion;
                RETURN;
            END;

            IF @TipoAyuda = 'TALLASPESOS'
            BEGIN
                SELECT T.IdTallaPeso AS cod, T.GlsTallaPeso AS des
                FROM dbo.TallaPeso AS T
                WHERE T.IdEmpresa = @IdEmpresa
                  AND (T.GlsTallaPeso LIKE '%' + @Texto + '%' OR T.IdTallaPeso LIKE '%' + @Texto + '%')
                ORDER BY T.IdTallaPeso;
                RETURN;
            END;

            IF @TipoAyuda = 'CANALES'
            BEGIN
                SELECT C.IdCanal AS cod, C.GlsCanal AS des
                FROM dbo.Canal AS C
                WHERE C.IdEmpresa = @IdEmpresa
                  AND (C.GlsCanal LIKE '%' + @Texto + '%' OR C.IdCanal LIKE '%' + @Texto + '%')
                ORDER BY C.IdCanal;
                RETURN;
            END;

            IF @TipoAyuda = 'TIPODOCUMENTOIDENTIDAD'
            BEGIN
                SELECT T.IdTipoDocIdentidad AS cod, T.GlsTipoDocIdentidad AS des
                FROM dbo.TiposDocIdentidad AS T
                WHERE T.IdTipoDocIdentidad LIKE '%' + @Texto + '%' OR T.GlsTipoDocIdentidad LIKE '%' + @Texto + '%'
                ORDER BY T.IdTipoDocIdentidad;
                RETURN;
            END;

            IF @TipoAyuda = 'CAJASUSUARIO'
            BEGIN
                SELECT C.IdCaja AS cod, C.GlsCaja AS des
                FROM dbo.Cajas AS C
                INNER JOIN dbo.CajasUsuario AS U
                    ON U.IdCaja = C.IdCaja
                   AND U.IdEmpresa = C.IdEmpresa
                WHERE C.IdEmpresa = @IdEmpresa
                  AND U.IdUsuario = @IdUsuario
                  AND (C.GlsCaja LIKE '%' + @Texto + '%' OR C.IdCaja LIKE '%' + @Texto + '%')
                ORDER BY C.IdCaja;
                RETURN;
            END;

            IF @TipoAyuda = 'ESTDOCUMENTOS'
            BEGIN
                SELECT DISTINCT
                    D.EstDocVentas AS cod,
                    CASE D.EstDocVentas
                        WHEN 'ANU' THEN 'ANULADO'
                        WHEN 'CAN' THEN 'CANCELADO'
                        WHEN 'GEN' THEN 'GENERADO'
                        WHEN 'IMP' THEN 'IMPRESO'
                    END AS des
                FROM dbo.DocVentas AS D
                WHERE D.EstDocVentas LIKE '%' + @Texto + '%'
                ORDER BY D.EstDocVentas;
                RETURN;
            END;

            RAISERROR ('Tipo de ayuda no implementado en spu_Ayuda_Catalogos.', 16, 1);
            RETURN;
        END;

        SELECT D.IdDato AS cod, D.GlsDato AS des
        FROM dbo.Datos AS D
        WHERE D.IdTipoDatos = @IdTipoDatos
          AND (D.GlsDato LIKE '%' + @Texto + '%' OR D.IdDato LIKE '%' + @Texto + '%')
        ORDER BY D.IdDato;
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
