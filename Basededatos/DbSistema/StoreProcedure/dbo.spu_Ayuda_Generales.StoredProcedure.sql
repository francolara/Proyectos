-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/10/2026
-- Description:   Atiende ayudas generales de FrmBusqueda, incluidos destinos, contactos y documentos de exportacion.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Ayuda_Generales
    @TipoAyuda VARCHAR(30),
    @Texto VARCHAR(200),
    @IdEmpresa CHAR(2),
    @IdSucursalContexto CHAR(8),
    @IdSucursalFiltro CHAR(8) = '',
    @IdPais VARCHAR(8) = '',
    @IdDepartamento VARCHAR(8) = '',
    @IdProvincia VARCHAR(8) = '',
    @IdDocumento VARCHAR(8) = '',
    @IndIngresoSalida VARCHAR(2) = '',
    @IdTipoNivel VARCHAR(8) = '',
    @IdNivelPred VARCHAR(20) = '',
    @IdProducto VARCHAR(20) = '',
    @IdCliente VARCHAR(20) = ''
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @TipoAyuda = UPPER(LTRIM(RTRIM(@TipoAyuda)));
        SET @Texto = LTRIM(RTRIM(ISNULL(@Texto, '')));

        IF @TipoAyuda = 'DEPARTAMENTO'
        BEGIN
            SELECT U.IdDpto AS cod, U.GlsUbigeo AS des
            FROM dbo.Ubigeo AS U
            WHERE U.IdProv = '00'
              AND U.IdDist = '00'
              AND (NULLIF(@IdPais, '') IS NULL OR U.IdPais = @IdPais)
              AND (U.GlsUbigeo LIKE '%' + @Texto + '%' OR U.IdDpto LIKE '%' + @Texto + '%')
            ORDER BY U.IdDpto;
            RETURN;
        END;

        IF @TipoAyuda = 'PROVINCIA'
        BEGIN
            SELECT U.IdProv AS cod, U.GlsUbigeo AS des
            FROM dbo.Ubigeo AS U
            WHERE U.IdProv <> '00'
              AND U.IdDist = '00'
              AND (NULLIF(@IdPais, '') IS NULL OR U.IdPais = @IdPais)
              AND (NULLIF(@IdDepartamento, '') IS NULL OR U.IdDpto = @IdDepartamento)
              AND (U.GlsUbigeo LIKE '%' + @Texto + '%' OR U.IdProv LIKE '%' + @Texto + '%')
            ORDER BY U.IdProv;
            RETURN;
        END;

        IF @TipoAyuda = 'DISTRITO'
        BEGIN
            SELECT U.IdDistrito AS cod, U.GlsUbigeo AS des
            FROM dbo.Ubigeo AS U
            WHERE U.IdDist <> '00'
              AND (NULLIF(@IdPais, '') IS NULL OR U.IdPais = @IdPais)
              AND (NULLIF(@IdDepartamento, '') IS NULL OR U.IdDpto = @IdDepartamento)
              AND (NULLIF(@IdProvincia, '') IS NULL OR U.IdProv = @IdProvincia)
              AND (U.GlsUbigeo LIKE '%' + @Texto + '%' OR U.IdDistrito LIKE '%' + @Texto + '%')
            ORDER BY U.IdDistrito;
            RETURN;
        END;

        IF @TipoAyuda = 'ALMACENVTA'
        BEGIN
            SELECT A.IdAlmacen AS cod, A.GlsAlmacen AS des
            FROM dbo.Almacenes AS A
            WHERE A.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdSucursalFiltro, '') IS NULL OR A.IdSucursal = @IdSucursalFiltro)
              AND (A.GlsAlmacen LIKE '%' + @Texto + '%' OR A.IdAlmacen LIKE '%' + @Texto + '%')
            ORDER BY A.IdAlmacen;
            RETURN;
        END;

        IF @TipoAyuda = 'MOTIVONCD'
        BEGIN
            SELECT M.IdMotivoNCD AS cod, M.GlsMotivoNCD AS des
            FROM dbo.MotivosNCD AS M
            WHERE (NULLIF(@IdDocumento, '') IS NULL OR M.IdDocumento = @IdDocumento)
              AND (M.GlsMotivoNCD LIKE '%' + @Texto + '%' OR M.IdMotivoNCD LIKE '%' + @Texto + '%')
            ORDER BY M.IdMotivoNCD;
            RETURN;
        END;

        IF @TipoAyuda = 'MOTIVOTRASLADO'
        BEGIN
            SELECT M.IdMotivoTraslado AS cod, M.GlsMotivoTraslado AS des
            FROM dbo.MotivosTraslados AS M
            WHERE (NULLIF(@IdDocumento, '') IS NULL OR M.IdDocumento = @IdDocumento)
              AND (M.GlsMotivoTraslado LIKE '%' + @Texto + '%' OR M.IdMotivoTraslado LIKE '%' + @Texto + '%')
            ORDER BY M.IdMotivoTraslado;
            RETURN;
        END;

        IF @TipoAyuda = 'TIPOMOVCAJA'
        BEGIN
            SELECT T.IdTipoMovCaja AS cod, T.GlsTipoMovCaja AS des
            FROM dbo.TiposMovCaja AS T
            WHERE LEFT(T.IdTipoMovCaja, 2) <> '99'
              AND (NULLIF(@IndIngresoSalida, '') IS NULL OR T.IndIngresoSalida = @IndIngresoSalida)
              AND (T.GlsTipoMovCaja LIKE '%' + @Texto + '%' OR T.IdTipoMovCaja LIKE '%' + @Texto + '%')
            ORDER BY T.IdTipoMovCaja;
            RETURN;
        END;

        IF @TipoAyuda = 'NIVEL'
        BEGIN
            SELECT N.IdNivel AS cod, N.GlsNivel AS des
            FROM dbo.Niveles AS N
            WHERE N.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdTipoNivel, '') IS NULL OR N.IdTipoNivel = @IdTipoNivel)
              AND (NULLIF(@IdNivelPred, '') IS NULL OR N.IdNivelPred = @IdNivelPred)
              AND (N.GlsNivel LIKE '%' + @Texto + '%' OR N.IdNivel LIKE '%' + @Texto + '%')
            ORDER BY N.IdNivel;
            RETURN;
        END;

        IF @TipoAyuda = 'PRESENTACIONES'
        BEGIN
            SELECT P.IdUM AS cod, U.AbreUM AS des
            FROM dbo.Presentaciones AS P
            INNER JOIN dbo.UnidadMedida AS U
                ON U.IdUM = P.IdUM
            WHERE P.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdProducto, '') IS NULL OR P.IdProducto = @IdProducto)
              AND (U.GlsUM LIKE '%' + @Texto + '%' OR P.IdUM LIKE '%' + @Texto + '%')
            ORDER BY P.IdUM;
            RETURN;
        END;

        IF @TipoAyuda = 'FORMASPAGOXCLIENTE'
        BEGIN
            SELECT C.IdFormaPago AS cod, F.GlsFormaPago AS des
            FROM dbo.ClientesFormaPagos AS C
            INNER JOIN dbo.FormasPagos AS F
                ON F.IdFormaPago = C.IdFormaPago
               AND F.IdEmpresa = C.IdEmpresa
            WHERE C.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdCliente, '') IS NULL OR C.IdCliente = @IdCliente)
              AND (F.GlsFormaPago LIKE '%' + @Texto + '%' OR C.IdFormaPago LIKE '%' + @Texto + '%')
            ORDER BY C.IdFormaPago;
            RETURN;
        END;

        IF @TipoAyuda = 'ALMACENDESTINO'
        BEGIN
            SELECT A.IdAlmacen AS cod, A.GlsAlmacen AS des
            FROM dbo.Almacenes AS A
            WHERE A.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdSucursalFiltro, '') IS NULL OR A.IdSucursal = @IdSucursalFiltro)
            ORDER BY A.IdAlmacen;
            RETURN;
        END;

        IF @TipoAyuda = 'SUCURSALDESTINO'
        BEGIN
            SELECT S.IdSucursal AS cod, S.GlsAbrev AS des
            FROM dbo.Sucursales AS S
            WHERE S.IdEmpresa = @IdEmpresa
            ORDER BY S.IdSucursal;
            RETURN;
        END;

        IF @TipoAyuda = 'DOCUMENOSEXP'
        BEGIN
            SELECT E.IdDocumentoExp AS cod, D.GlsDocumento AS des
            FROM dbo.DocumentosExportar AS E
            INNER JOIN dbo.Documentos AS D
                ON D.IdDocumento = E.IdDocumentoExp
            WHERE (NULLIF(@IdDocumento, '') IS NULL OR E.IdDocumento = @IdDocumento)
              AND (D.GlsDocumento LIKE '%' + @Texto + '%' OR E.IdDocumentoExp LIKE '%' + @Texto + '%')
            ORDER BY E.IdDocumentoExp;
            RETURN;
        END;

        IF @TipoAyuda = 'CONTACTOSCLIENTES'
        BEGIN
            SELECT C.IdContacto AS cod, P.GlsPersona AS des
            FROM dbo.ContactosClientes AS C
            INNER JOIN dbo.Personas AS P
                ON P.IdPersona = C.IdContacto
            WHERE C.IdEmpresa = @IdEmpresa
              AND (NULLIF(@IdCliente, '') IS NULL OR C.IdCliente = @IdCliente)
              AND (P.GlsContacto LIKE '%' + @Texto + '%' OR C.IdContacto LIKE '%' + @Texto + '%')
            ORDER BY C.IdContacto;
            RETURN;
        END;

        IF @TipoAyuda = 'CONTACTOSPROVEEDORES'
        BEGIN
            SELECT C.IdContacto AS cod, P.GlsPersona AS des
            FROM dbo.ContactosProveedores AS C
            INNER JOIN dbo.Personas AS P
                ON P.IdPersona = C.IdContacto
            WHERE C.IdEmpresa = @IdEmpresa
              AND (P.GlsContacto LIKE '%' + @Texto + '%' OR C.IdContacto LIKE '%' + @Texto + '%')
            ORDER BY C.IdContacto;
            RETURN;
        END;

        RAISERROR ('Tipo de ayuda no implementado en spu_Ayuda_Generales.', 16, 1);
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
