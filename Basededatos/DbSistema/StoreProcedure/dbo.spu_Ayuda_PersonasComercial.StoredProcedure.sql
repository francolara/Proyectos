-- =============================================
-- Author:        FRANCO LARA
-- Create date:   15/10/2026
-- Description:   Atiende ayudas frecuentes de personas, comerciales y seguridad para FrmBusqueda.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_Ayuda_PersonasComercial
    @TipoAyuda VARCHAR(30),
    @Texto VARCHAR(200),
    @IdEmpresa CHAR(2),
    @IdSucursal CHAR(8),
    @IdUsuario CHAR(8),
    @FiltrarClientePorVendedor BIT = 0
AS
BEGIN
    SET NOCOUNT ON;

    BEGIN TRY
        SET @TipoAyuda = UPPER(LTRIM(RTRIM(@TipoAyuda)));
        SET @Texto = LTRIM(RTRIM(ISNULL(@Texto, '')));

        IF @TipoAyuda = 'CLIENTE'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Clientes AS C
                ON C.IdCliente = P.IdPersona
               AND C.IdEmpresa = @IdEmpresa
            WHERE (@FiltrarClientePorVendedor = 0 OR C.IdVendedorCampo = @IdUsuario)
              AND (P.GlsPersona LIKE '%' + @Texto + '%' OR P.IdPersona LIKE '%' + @Texto + '%')
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'PROVEEDOR'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Proveedores AS PR
                ON PR.IdProveedor = P.IdPersona
               AND PR.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'CAMAL'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Proveedores AS PR
                ON PR.IdProveedor = P.IdPersona
               AND PR.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'VENDEDOR'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Vendedores AS V
                ON V.IdVendedor = P.IdPersona
               AND V.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'VENDEDORJEFE'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Vendedores AS V
                ON V.IdVendedor = P.IdPersona
               AND V.IdEmpresa = @IdEmpresa
            WHERE V.IndJefe = '1'
              AND (P.GlsPersona LIKE '%' + @Texto + '%' OR P.IdPersona LIKE '%' + @Texto + '%')
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda IN ('USUARIO', 'USUARIOS')
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Usuarios AS U
                ON U.IdUsuario = P.IdPersona
               AND U.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'CHOFER'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Choferes AS CH
                ON CH.IdChofer = P.IdPersona
               AND CH.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'EMPTRANS'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.EmpTrans AS ET
                ON ET.IdEmpTrans = P.IdPersona
               AND ET.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'COBRADOR'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Cobradores AS C
                ON C.IdCobrador = P.IdPersona
               AND C.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'AVAL'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Aval AS A
                ON A.IdAval = P.IdPersona
               AND A.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'SUCURSAL'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            INNER JOIN dbo.Sucursales AS S
                ON S.IdSucursal = P.IdPersona
               AND S.IdEmpresa = @IdEmpresa
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'PERSONA'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            WHERE P.GlsPersona LIKE '%' + @Texto + '%'
               OR P.IdPersona LIKE '%' + @Texto + '%'
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'PERSONACLIENTE'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            WHERE NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Clientes AS C
                    WHERE C.IdEmpresa = @IdEmpresa
                      AND C.IdCliente = P.IdPersona
                )
              AND (P.GlsPersona LIKE '%' + @Texto + '%' OR P.IdPersona LIKE '%' + @Texto + '%')
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'PERSONASUCURSAL'
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            WHERE P.TipoPersona = '01002'
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Sucursales AS S
                    WHERE S.IdEmpresa = @IdEmpresa
                      AND S.IdSucursal = P.IdPersona
                )
              AND (P.GlsPersona LIKE '%' + @Texto + '%' OR P.IdPersona LIKE '%' + @Texto + '%')
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda IN ('PERSONAPROVEEDOR', 'PERSONAVENDEDOR', 'PERSONAUSUARIO', 'PERSONACHOFER', 'PERSONAEMPTRANS', 'PERSONACOBRADOR', 'PERSONAAVAL')
        BEGIN
            SELECT
                P.IdPersona AS cod,
                P.GlsPersona AS des
            FROM dbo.Personas AS P
            WHERE (@TipoAyuda NOT IN ('PERSONAVENDEDOR', 'PERSONAUSUARIO', 'PERSONACHOFER') OR P.TipoPersona = '01001')
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Proveedores AS PR
                    WHERE @TipoAyuda = 'PERSONAPROVEEDOR'
                      AND PR.IdEmpresa = @IdEmpresa
                      AND PR.IdProveedor = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Vendedores AS V
                    WHERE @TipoAyuda = 'PERSONAVENDEDOR'
                      AND V.IdEmpresa = @IdEmpresa
                      AND V.IdVendedor = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Usuarios AS U
                    WHERE @TipoAyuda = 'PERSONAUSUARIO'
                      AND U.IdEmpresa = @IdEmpresa
                      AND U.IdUsuario = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Choferes AS CH
                    WHERE @TipoAyuda = 'PERSONACHOFER'
                      AND CH.IdEmpresa = @IdEmpresa
                      AND CH.IdChofer = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.EmpTrans AS ET
                    WHERE @TipoAyuda = 'PERSONAEMPTRANS'
                      AND ET.IdEmpresa = @IdEmpresa
                      AND ET.IdEmpTrans = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Cobradores AS C
                    WHERE @TipoAyuda = 'PERSONACOBRADOR'
                      AND C.IdEmpresa = @IdEmpresa
                      AND C.IdCobrador = P.IdPersona
                )
              AND NOT EXISTS
                (
                    SELECT 1
                    FROM dbo.Aval AS A
                    WHERE @TipoAyuda = 'PERSONAAVAL'
                      AND A.IdEmpresa = @IdEmpresa
                      AND A.IdAval = P.IdPersona
                )
              AND (P.GlsPersona LIKE '%' + @Texto + '%' OR P.IdPersona LIKE '%' + @Texto + '%')
            ORDER BY P.IdPersona;
            RETURN;
        END;

        IF @TipoAyuda = 'PRODUCTOS'
        BEGIN
            SELECT
                P.IdProducto AS cod,
                P.GlsProducto AS des
            FROM dbo.Productos AS P
            WHERE P.IdEmpresa = @IdEmpresa
              AND P.EstProducto = 'A'
              AND (P.GlsProducto LIKE '%' + @Texto + '%' OR P.IdProducto LIKE '%' + @Texto + '%')
            ORDER BY P.IdProducto;
            RETURN;
        END;

        IF @TipoAyuda = 'ALMACEN'
        BEGIN
            SELECT
                A.IdAlmacen AS cod,
                A.GlsAlmacen AS des
            FROM dbo.Almacenes AS A
            WHERE A.IdEmpresa = @IdEmpresa
              AND A.IdSucursal = @IdSucursal
              AND (A.GlsAlmacen LIKE '%' + @Texto + '%' OR A.IdAlmacen LIKE '%' + @Texto + '%')
            ORDER BY A.IdAlmacen;
            RETURN;
        END;

        IF @TipoAyuda = 'ALMACENVTA'
        BEGIN
            SELECT
                A.IdAlmacen AS cod,
                A.GlsAlmacen AS des
            FROM dbo.Almacenes AS A
            WHERE A.IdEmpresa = @IdEmpresa
              AND (A.GlsAlmacen LIKE '%' + @Texto + '%' OR A.IdAlmacen LIKE '%' + @Texto + '%')
            ORDER BY A.IdAlmacen;
            RETURN;
        END;

        IF @TipoAyuda = 'CENTROCOSTO'
        BEGIN
            SELECT
                CC.IdCentroCosto AS cod,
                CC.GlsCentroCosto AS des
            FROM dbo.CentrosCosto AS CC
            WHERE CC.IdEmpresa = @IdEmpresa
              AND (CC.GlsCentroCosto LIKE '%' + @Texto + '%' OR CC.IdCentroCosto LIKE '%' + @Texto + '%')
            ORDER BY CC.IdCentroCosto;
            RETURN;
        END;

        IF @TipoAyuda = 'LISTAPRECIOS'
        BEGIN
            SELECT
                LP.IdLista AS cod,
                LP.GlsLista AS des
            FROM dbo.ListaPrecios AS LP
            WHERE LP.IdEmpresa = @IdEmpresa
              AND LP.EstLista = 1
              AND (LP.GlsLista LIKE '%' + @Texto + '%' OR LP.IdLista LIKE '%' + @Texto + '%')
            ORDER BY LP.IdLista;
            RETURN;
        END;

        IF @TipoAyuda = 'VEHICULO'
        BEGIN
            SELECT
                V.IdVehiculo AS cod,
                V.GlsVehiculo AS des
            FROM dbo.Vehiculos AS V
            WHERE V.IdEmpresa = @IdEmpresa
              AND (V.GlsVehiculo LIKE '%' + @Texto + '%' OR V.IdVehiculo LIKE '%' + @Texto + '%')
            ORDER BY V.IdVehiculo;
            RETURN;
        END;

        RAISERROR ('Tipo de ayuda no implementado en spu_Ayuda_PersonasComercial.', 16, 1);
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
