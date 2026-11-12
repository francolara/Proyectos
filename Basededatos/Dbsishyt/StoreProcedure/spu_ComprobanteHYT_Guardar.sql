-- Firma: FRANCO LARA | Create date: 10/11/2026 | Operaciones de provisiones y ventas. Admite carga atómica y alta de clientes desde XML y conserva la fecha de vencimiento y calcula la detracción manual con el catálogo de FralseCont y admite depósito en otra moneda con equivalente calculado.
-- Firma: FRANCO LARA | Create date: 11/11/2026 | Calcula la detracción de ventas por materiales o conserva los datos declarados en XML y vincula la orden de compra del cliente y valida la factura y el motivo SUNAT en notas de crédito manuales de venta.
USE [SistemaWebHYT];
GO
CREATE OR ALTER PROCEDURE dbo.spu_ComprobanteHYT_Guardar
    @UsuarioId NVARCHAR(128), @EsVenta BIT, @Datos NVARCHAR(MAX), @Id BIGINT=NULL
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;
    BEGIN TRY
        IF dbo.fn_HYTPermiso(@UsuarioId,CASE WHEN @EsVenta=1 THEN N'FACTURADO.EDITAR' ELSE N'PAGOS_PROVEEDOR.EDITAR' END)=0
            THROW 51600,N'[HYT_PERMISO] No tiene permiso para esta operación.',1;
        IF ISJSON(@Datos)<>1 OR @Datos IS NULL
            THROW 51601,N'[HYT_VALIDACION] Datos del comprobante inválidos.',1;
        DECLARE @RucHYT NVARCHAR(20),@Tercero NVARCHAR(30),@Nombre NVARCHAR(180),@Tipo CHAR(2),@Serie NVARCHAR(30),@Numero NVARCHAR(50),
            @Fecha DATE,@FechaVencimiento DATE,@Moneda NVARCHAR(3),@Forma NVARCHAR(30),@Importe DECIMAL(18,2),@Detalle NVARCHAR(MAX),@Xml NVARCHAR(MAX),
            @Detraccion DECIMAL(18,2),@MonedaDetraccion NVARCHAR(3),@CodigoDetraccion NVARCHAR(30),@Ref NVARCHAR(100),@Motivo NVARCHAR(300),@CodigoMotivo NVARCHAR(4000),
            @ClienteId BIGINT,@ProveedorId BIGINT,@Presupuestos NVARCHAR(MAX),@RucDocumento NVARCHAR(20),@DetraccionDocumento DECIMAL(18,2),@PorcentajeDetraccion DECIMAL(7,4),@OrdenCompraId BIGINT;
        SELECT @Tipo=Tipo,@Serie=NULLIF(Serie,N''),@Numero=NULLIF(Numero,N''),@Fecha=Fecha,@FechaVencimiento=FechaVencimiento,@Moneda=Moneda,@Forma=FormaPago,@Importe=Importe,
            @Detalle=Detalle,@Xml=XmlOriginal,@Detraccion=ImporteDetraccion,@MonedaDetraccion=MonedaDetraccion,@CodigoDetraccion=NULLIF(TRIM(CodigoDetraccion),N''),@PorcentajeDetraccion=PorcentajeDetraccion,@OrdenCompraId=OrdenCompraId,
            @Ref=ReferenciaNota,@Motivo=MotivoNota,@CodigoMotivo=CodigoMotivoNota,@Tercero=DocumentoTercero,@Nombre=NombreTercero,@Presupuestos=Presupuestos,@RucDocumento=RucHYT
        FROM OPENJSON(@Datos) WITH (
            Tipo CHAR(2),Serie NVARCHAR(30),Numero NVARCHAR(50),Fecha DATE,FechaVencimiento DATE,Moneda NVARCHAR(3),FormaPago NVARCHAR(30),Importe DECIMAL(18,2),
            Detalle NVARCHAR(MAX) AS JSON,XmlOriginal NVARCHAR(MAX),ImporteDetraccion DECIMAL(18,2),MonedaDetraccion NVARCHAR(3),CodigoDetraccion NVARCHAR(30),PorcentajeDetraccion DECIMAL(7,4),OrdenCompraId BIGINT,
            ReferenciaNota NVARCHAR(100),MotivoNota NVARCHAR(300),CodigoMotivoNota NVARCHAR(4000),DocumentoTercero NVARCHAR(30),NombreTercero NVARCHAR(180),
            Presupuestos NVARCHAR(MAX) AS JSON,RucHYT NVARCHAR(20));
        IF @CodigoMotivo IS NOT NULL AND (LEN(@CodigoMotivo)<>2 OR @CodigoMotivo COLLATE Latin1_General_100_BIN2 LIKE N'%[^0-9]%')
            THROW 51601,N'[HYT_VALIDACION] El código de motivo de la nota de crédito debe tener dos dígitos.',1;
        IF @EsVenta=1 AND @Tipo='07' AND @Xml IS NULL
        BEGIN
            SET @Motivo=NULL;
            SELECT @Motivo=motivo.Descripcion FROM dbo.MotivoNotaCreditoHYT AS motivo WHERE motivo.Codigo=@CodigoMotivo;
            IF @Motivo IS NULL
                THROW 51601,N'[HYT_VALIDACION] Selecciona un motivo SUNAT para la nota de crédito.',1;
        END;
        IF @Xml IS NULL
        BEGIN
            SET @PorcentajeDetraccion=NULL;
            IF @EsVenta=1 AND @Tipo='01' AND ISJSON(@Detalle,ARRAY)=1
                SELECT @CodigoDetraccion=material.Codigo,@PorcentajeDetraccion=material.Porcentaje
                FROM dbo.fn_ComprobanteHYT_DetraccionMaterial(@Detalle) AS material;
            IF @EsVenta=1 AND @Tipo<>'01' SET @CodigoDetraccion=NULL;

            IF @CodigoDetraccion IS NOT NULL AND @PorcentajeDetraccion IS NULL
            BEGIN
                SET @PorcentajeDetraccion=NULL;
                SELECT @PorcentajeDetraccion=catalogo.Porcentaje FROM dbo.DetraccionCatalogoHYT AS catalogo
                    WHERE catalogo.Codigo=@CodigoDetraccion AND catalogo.Activo=1;
                IF @PorcentajeDetraccion IS NULL
                    THROW 51601,N'[HYT_VALIDACION] Selecciona un código de detracción activo del catálogo.',1;
            END;
            SET @PorcentajeDetraccion=COALESCE(@PorcentajeDetraccion,0);
            SET @DetraccionDocumento=ROUND(@Importe*@PorcentajeDetraccion/100,2);
            IF @EsVenta=1 OR @CodigoDetraccion IS NULL OR @MonedaDetraccion IS NULL OR @MonedaDetraccion=@Moneda
            BEGIN
                SET @Detraccion=@DetraccionDocumento;
                SET @MonedaDetraccion=@Moneda;
            END
            ELSE IF COALESCE(@Detraccion,0)<=0
                THROW 51601,N'[HYT_VALIDACION] Indica el importe a depositar en la moneda de la detracción.',1;
        END;
        ELSE
        BEGIN
            SET @DetraccionDocumento=COALESCE(TRY_CONVERT(DECIMAL(18,2),JSON_VALUE(@Datos,N'$.DetraccionDocumento')),
                CASE WHEN @Detraccion=0 THEN 0 WHEN @MonedaDetraccion=@Moneda THEN @Detraccion END);
            IF @EsVenta=1
            BEGIN
                IF @PorcentajeDetraccion IS NULL AND @Tipo='01'
                    SELECT @PorcentajeDetraccion=catalogo.Porcentaje FROM dbo.DetraccionCatalogoHYT AS catalogo
                    WHERE catalogo.Codigo=@CodigoDetraccion;
                SET @PorcentajeDetraccion=COALESCE(@PorcentajeDetraccion,
                    TRY_CONVERT(DECIMAL(7,4),100*@DetraccionDocumento/NULLIF(@Importe,0)),0);
                IF @Tipo='01' AND @Detraccion=0 AND @PorcentajeDetraccion>0
                BEGIN
                    SET @DetraccionDocumento=ROUND(@Importe*@PorcentajeDetraccion/100,2);
                    SET @Detraccion=@DetraccionDocumento;
                    SET @MonedaDetraccion=@Moneda;
                END;
            END;
        END;
        IF @Tipo IS NULL OR @Tipo NOT IN ('01','03','07') OR @Fecha IS NULL OR @Moneda IS NULL OR @Moneda NOT IN(N'PEN',N'USD')
            OR @Forma IS NULL OR @Forma NOT IN(N'Contado',N'Crédito',N'No indicada')
            OR @Tercero IS NULL OR LEN(@Tercero)=0 OR (@EsVenta=0 AND (LEN(@Tercero)<>11 OR @Tercero LIKE N'%[^0-9]%'))
            OR (@EsVenta=0 AND NULLIF(@Nombre,N'') IS NULL)
            OR @Importe IS NULL OR @Importe<0 OR @Detalle IS NULL OR NOT EXISTS(SELECT 1 FROM OPENJSON(@Detalle))
            OR (@Tipo='07' AND (NULLIF(@Ref,N'') IS NULL OR NULLIF(@Motivo,N'') IS NULL))
            OR ((@EsVenta=0 OR @Xml IS NOT NULL) AND (@Serie IS NULL OR @Numero IS NULL))
            OR (@Serie IS NULL AND @Numero IS NOT NULL) OR (@Numero IS NULL AND @Serie IS NOT NULL)
            OR @Detraccion IS NULL OR @Detraccion<0 OR @MonedaDetraccion IS NULL OR @MonedaDetraccion NOT IN(N'PEN',N'USD')
            OR @PorcentajeDetraccion NOT BETWEEN 0 AND 100
            OR @DetraccionDocumento IS NULL OR @DetraccionDocumento<0 OR @DetraccionDocumento>@Importe
            OR (@Detraccion>0 AND @DetraccionDocumento<=0)
            OR (@MonedaDetraccion=@Moneda AND @DetraccionDocumento<>@Detraccion)
            THROW 51601,N'[HYT_VALIDACION] Revisa los datos y el detalle del comprobante.',1;
        IF (SELECT COUNT(1) FROM OPENJSON(@Detalle))>500 OR EXISTS (
            SELECT 1 FROM OPENJSON(@Detalle) WITH(Descripcion NVARCHAR(MAX),Unidad NVARCHAR(MAX),Cantidad DECIMAL(19,6),Precio DECIMAL(19,6),Importe DECIMAL(19,6)) AS linea
            WHERE NULLIF(LTRIM(RTRIM(linea.Descripcion)),N'') IS NULL OR LEN(linea.Descripcion)>1500
                OR NULLIF(LTRIM(RTRIM(linea.Unidad)),N'') IS NULL OR LEN(linea.Unidad)>30
                OR linea.Cantidad IS NULL OR linea.Cantidad<0 OR linea.Precio IS NULL OR linea.Precio<0 OR linea.Importe IS NULL OR linea.Importe<0)
            THROW 51601,N'[HYT_VALIDACION] Revisa las partidas del comprobante.',1;
        DECLARE @OperacionRegistro UNIQUEIDENTIFIER=COALESCE(TRY_CONVERT(UNIQUEIDENTIFIER,JSON_VALUE(@Datos,N'$.OperacionId')),NEWID()),
            @Huella VARBINARY(32)=HASHBYTES('SHA2_256',@Datos),@Existente BIGINT;
        BEGIN TRANSACTION;
        IF @Id IS NULL
        BEGIN
            SELECT @Existente=ComprobanteId FROM dbo.ComprobanteHYT WITH(UPDLOCK,HOLDLOCK) WHERE OperacionId=@OperacionRegistro;
            IF @Existente IS NOT NULL
            BEGIN
                IF NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYT WHERE ComprobanteId=@Existente AND Huella=@Huella AND UsuarioRegistro=@UsuarioId AND EsVenta=@EsVenta)
                    THROW 51601,N'[HYT_VALIDACION] La operación ya se utilizó con otros datos.',1;
                COMMIT TRANSACTION;
                SELECT @Existente;
                RETURN;
            END
        END;
        SELECT @RucHYT=Ruc FROM dbo.ConfiguracionHYT WITH(HOLDLOCK) WHERE Id=1;
        IF @RucHYT IS NULL OR LEN(@RucHYT)<>11 OR @RucHYT LIKE N'%[^0-9]%' OR @Tercero=@RucHYT OR (@Xml IS NOT NULL AND (@RucDocumento IS NULL OR @RucDocumento<>@RucHYT))
            THROW 51601,N'[HYT_VALIDACION] El RUC de HYT debe coincidir con la configuración y figurar como emisor de ventas o receptor de compras.',1;
        IF @EsVenta=1
        BEGIN
            IF @Xml IS NOT NULL AND @Id IS NULL
            BEGIN
                DECLARE @BloqueoCliente INT;
                EXEC @BloqueoCliente=sys.sp_getapplock @Resource=N'HYT:Codigo:Cliente',@LockMode=N'Exclusive',@LockOwner=N'Transaction',@LockTimeout=10000;
                IF @BloqueoCliente<0 THROW 51601,N'[HYT_VALIDACION] No se pudo consultar el cliente. Inténtalo nuevamente.',1;
            END;
            SELECT TOP (1) @ClienteId=cliente.ClienteId,
                @Nombre=COALESCE(NULLIF(@Nombre,N''),CASE WHEN cliente.TipoEntidad=N'NATURAL' THEN CONCAT_WS(N' ',cliente.Nombres,cliente.Apellidos) ELSE cliente.RazonSocial END)
            FROM dbo.Cliente AS cliente WITH(HOLDLOCK) WHERE cliente.NumeroIdentificacion=@Tercero AND cliente.Activo=1 ORDER BY cliente.ClienteId;
            IF @ClienteId IS NULL AND @Xml IS NOT NULL AND @Id IS NULL AND NOT EXISTS(
                SELECT 1 FROM dbo.Cliente WHERE NumeroIdentificacion=@Tercero)
            BEGIN
                IF NULLIF(LTRIM(RTRIM(@Nombre)),N'') IS NULL
                    THROW 51601,N'[HYT_VALIDACION] El XML debe identificar el nombre del cliente nuevo.',1;
                DECLARE @Codigo NVARCHAR(40),@Secuencia NVARCHAR(20),@TipoIdentificacion NVARCHAR(20),
                    @Juridica BIT=0,@Separador INT=CHARINDEX(N' ',@Nombre);
                IF @Tercero NOT LIKE N'%[^0-9]%' AND LEN(@Tercero)=11
                BEGIN
                    SET @TipoIdentificacion=N'RUC';
                    IF LEFT(@Tercero,1)<>N'1' SET @Juridica=1;
                END;
                IF @Tercero NOT LIKE N'%[^0-9]%' AND LEN(@Tercero)=8 SET @TipoIdentificacion=N'DNI';
                IF @Juridica=0 AND (@Separador<2 OR @Separador>151 OR LEN(LTRIM(SUBSTRING(@Nombre,@Separador+1,180))) NOT BETWEEN 1 AND 150)
                    THROW 51601,N'[HYT_VALIDACION] El XML debe contener el nombre completo del cliente nuevo.',1;
                WHILE 1=1
                BEGIN
                    SET @Secuencia=CONVERT(NVARCHAR(20),NEXT VALUE FOR dbo.SecuenciaCodigoInterno);
                    SET @Codigo=N'CLI-'+REPLICATE(N'0',CASE WHEN LEN(@Secuencia)<6 THEN 6-LEN(@Secuencia) ELSE 0 END)+@Secuencia;
                    IF NOT EXISTS(SELECT 1 FROM dbo.Cliente WHERE Codigo=@Codigo) BREAK;
                END;
                INSERT INTO dbo.Cliente(Codigo,TipoEntidad,TipoIdentificacion,NumeroIdentificacion,RazonSocial,Nombres,Apellidos,UsuarioRegistro)
                VALUES(@Codigo,CASE WHEN @Juridica=1 THEN N'JURIDICA' ELSE N'NATURAL' END,@TipoIdentificacion,@Tercero,
                    CASE WHEN @Juridica=1 THEN @Nombre END,CASE WHEN @Juridica=0 THEN LEFT(@Nombre,@Separador-1) END,
                    CASE WHEN @Juridica=0 THEN LTRIM(SUBSTRING(@Nombre,@Separador+1,180)) END,@UsuarioId);
                SET @ClienteId=SCOPE_IDENTITY();
                DECLARE @RegistroCliente NVARCHAR(50)=CONVERT(NVARCHAR(50),@ClienteId),@OperacionCliente UNIQUEIDENTIFIER=NEWID();
                EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@OperacionCliente,N'Cliente',@RegistroCliente,N'CREAR';
            END;
            IF @ClienteId IS NULL THROW 51601,N'[HYT_VALIDACION] El cliente del comprobante debe existir y estar activo.',1;
            IF EXISTS(SELECT 1 FROM OPENJSON(@Presupuestos) AS seleccionado LEFT JOIN dbo.Presupuesto AS p ON p.PresupuestoId=TRY_CONVERT(BIGINT,seleccionado.value)
                WHERE p.PresupuestoId IS NULL OR p.ClienteId<>@ClienteId)
                THROW 51601,N'[HYT_VALIDACION] Los presupuestos deben pertenecer al cliente del comprobante.',1;
        END
        ELSE
        BEGIN
            SELECT @ProveedorId=ProveedorId FROM dbo.ProveedorHYT WITH(UPDLOCK,HOLDLOCK) WHERE Ruc=CONVERT(VARCHAR(11),@Tercero);
            IF @ProveedorId IS NULL
            BEGIN
                INSERT INTO dbo.ProveedorHYT(Ruc,RazonSocial) VALUES(@Tercero,@Nombre);
                SET @ProveedorId=SCOPE_IDENTITY();
            END
        END;
        IF @EsVenta=1 AND @Tipo='07' AND @Xml IS NULL AND NOT EXISTS (
            SELECT 1 FROM dbo.ComprobanteHYT AS factura WITH(HOLDLOCK)
            WHERE factura.EsVenta=1 AND factura.Tipo='01' AND factura.ClienteId=@ClienteId
                AND factura.Serie IS NOT NULL AND factura.Numero IS NOT NULL
                AND factura.EstadoSunat NOT IN(N'RECHAZADO',N'ANULADO',N'BAJA')
                AND CONCAT(factura.Serie,N'-',factura.Numero)=@Ref)
            THROW 51601,N'[HYT_VALIDACION] Selecciona una factura emitida del cliente.',1;

        IF @OrdenCompraId IS NOT NULL AND (@EsVenta=0 OR NOT EXISTS (
            SELECT 1 FROM dbo.OrdenCompraCliente AS orden WITH(HOLDLOCK)
            INNER JOIN dbo.Presupuesto AS p ON p.PresupuestoId=orden.PresupuestoId
            WHERE orden.OrdenCompraId=@OrdenCompraId AND p.ClienteId=@ClienteId AND orden.Estado<>N'CANCELADA'))
            THROW 51601,N'[HYT_VALIDACION] Selecciona una orden de compra vigente del cliente.',1;

        IF @Serie IS NOT NULL AND EXISTS(SELECT 1 FROM dbo.ComprobanteHYT WITH(UPDLOCK,HOLDLOCK)
            WHERE EsVenta=@EsVenta AND RucHYT=@RucHYT AND DocumentoTercero=@Tercero AND Tipo=@Tipo AND Serie=@Serie AND Numero=@Numero
            AND (@Id IS NULL OR ComprobanteId<>@Id))
            THROW 51601,N'[HYT_VALIDACION] El comprobante ya se registró.',1;
        IF @Id IS NULL
        BEGIN
            INSERT INTO dbo.ComprobanteHYT(OperacionId,Huella,EsVenta,Tipo,Serie,Numero,Fecha,FechaVencimiento,Moneda,RucHYT,DocumentoTercero,NombreTercero,ProveedorId,ClienteId,OrdenCompraId,FormaPago,Importe,
                Detalle,ImporteDetraccion,DetraccionDocumento,MonedaDetraccion,CodigoDetraccion,PorcentajeDetraccion,ReferenciaNota,MotivoNota,CodigoMotivoNota,XmlOriginal,UsuarioRegistro)
            VALUES(@OperacionRegistro,@Huella,@EsVenta,@Tipo,@Serie,@Numero,@Fecha,@FechaVencimiento,@Moneda,@RucHYT,@Tercero,@Nombre,@ProveedorId,@ClienteId,@OrdenCompraId,@Forma,@Importe,
                @Detalle,@Detraccion,@DetraccionDocumento,@MonedaDetraccion,@CodigoDetraccion,@PorcentajeDetraccion,@Ref,@Motivo,@CodigoMotivo,@Xml,@UsuarioId);
            SET @Id=SCOPE_IDENTITY();
        END
        ELSE
        BEGIN
            UPDATE dbo.ComprobanteHYT SET Tipo=@Tipo,Serie=@Serie,Numero=@Numero,Fecha=@Fecha,FechaVencimiento=COALESCE(@FechaVencimiento,FechaVencimiento),Moneda=@Moneda,FormaPago=@Forma,Importe=@Importe,
                Detalle=@Detalle,ReferenciaNota=@Ref,MotivoNota=@Motivo,CodigoMotivoNota=COALESCE(@CodigoMotivo,CodigoMotivoNota),XmlOriginal=@Xml,CodigoDetraccion=@CodigoDetraccion,
                OrdenCompraId=COALESCE(@OrdenCompraId,OrdenCompraId),
                PorcentajeDetraccion=@PorcentajeDetraccion,ImporteDetraccion=@Detraccion,DetraccionDocumento=@DetraccionDocumento,MonedaDetraccion=@MonedaDetraccion
            WHERE ComprobanteId=@Id AND EsVenta=1 AND @EsVenta=1 AND XmlOriginal IS NULL AND Serie IS NULL AND @Xml IS NOT NULL
                AND ClienteId=@ClienteId AND RucHYT=@RucHYT AND Tipo=@Tipo AND Moneda=@Moneda;
            IF @@ROWCOUNT=0 THROW 51601,N'[HYT_VALIDACION] Solo puedes completar un borrador con XML del mismo cliente, tipo y moneda.',1;
        END;
        IF @EsVenta=1
            INSERT INTO dbo.ComprobanteHYTPresupuesto(ComprobanteId,PresupuestoId)
            SELECT DISTINCT @Id,CONVERT(BIGINT,seleccionado.value) FROM OPENJSON(@Presupuestos) AS seleccionado
            WHERE NOT EXISTS(SELECT 1 FROM dbo.ComprobanteHYTPresupuesto AS vinculo WHERE vinculo.ComprobanteId=@Id AND vinculo.PresupuestoId=TRY_CONVERT(BIGINT,seleccionado.value));
        DECLARE @Registro NVARCHAR(50)=CONVERT(NVARCHAR(50),@Id),@Operacion UNIQUEIDENTIFIER=NEWID();
        EXEC dbo.spu_RegistrarAuditoria @UsuarioId,@Operacion,N'ComprobanteHYT',@Registro,N'REGISTRAR';
        COMMIT TRANSACTION;
        SELECT @Id;
    END TRY
    BEGIN CATCH
        IF XACT_STATE()<>0 ROLLBACK TRANSACTION;
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage,@ErrorSeverity,@ErrorState);
    END CATCH
END;
GO
GRANT EXECUTE ON OBJECT::dbo.spu_ComprobanteHYT_Guardar TO hyt_app_executor;
GO
