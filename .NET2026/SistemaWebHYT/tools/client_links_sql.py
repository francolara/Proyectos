"""Fragmentos estáticos del guardado transaccional de enlaces del cliente."""

SNAPSHOT = ", JSON_QUERY(COALESCE((SELECT enlace.Id,enlace.Descripcion,enlace.Url,enlace.Orden,enlace.FechaRegistro,enlace.UsuarioRegistro,enlace.FechaActualizacion,enlace.UsuarioActualizacion FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=anterior.ClienteId ORDER BY enlace.Orden,enlace.Id FOR JSON PATH,INCLUDE_NULL_VALUES),N'[]')) AS Enlaces"

VALIDATE = '''        SET @UrlCargaDocumentos=JSON_VALUE(@Anterior,N'$.UrlCargaDocumentos');
        DECLARE @Enlaces NVARCHAR(MAX)=JSON_QUERY(@Datos,N'$.Enlaces');
        IF EXISTS(SELECT 1 FROM OPENJSON(@Datos) AS dato WHERE dato.[key]=N'Enlaces') AND (@Enlaces IS NULL OR LEFT(LTRIM(@Enlaces),1)<>N'[')
            THROW 51226,N'[HYT_VALIDACION] La lista de enlaces no es válida.',1;
        DECLARE @ListaEnlaces TABLE(Id BIGINT NULL,Descripcion NVARCHAR(MAX),Url NVARCHAR(MAX),Orden INT);
        IF @Enlaces IS NOT NULL
        BEGIN
            IF (SELECT COUNT_BIG(1) FROM OPENJSON(@Enlaces))>100 OR EXISTS(SELECT 1 FROM OPENJSON(@Enlaces) AS jsonEnlace WHERE jsonEnlace.type<>5)
                THROW 51226,N'[HYT_VALIDACION] Revisa la lista de enlaces (máximo cien).',1;
            INSERT INTO @ListaEnlaces(Id,Descripcion,Url,Orden)
            SELECT datos.Id,NULLIF(LTRIM(RTRIM(datos.Descripcion)),N''),NULLIF(LTRIM(RTRIM(datos.Url)),N''),CONVERT(INT,entrada.[key])
            FROM OPENJSON(@Enlaces) AS entrada
            CROSS APPLY OPENJSON(entrada.value) WITH(Id BIGINT '$.Id',Descripcion NVARCHAR(MAX) '$.Descripcion',Url NVARCHAR(MAX) '$.Url') AS datos;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS enlace WHERE enlace.Descripcion IS NULL OR LEN(enlace.Descripcion)>200 OR enlace.Url IS NULL OR LEN(enlace.Url)>2048
                OR (enlace.Url NOT LIKE N'https://%' AND enlace.Url NOT LIKE N'http://%') OR CHARINDEX(NCHAR(10),enlace.Url)>0 OR CHARINDEX(NCHAR(13),enlace.Url)>0)
                THROW 51226,N'[HYT_VALIDACION] Completa descripción y enlace HTTP o HTTPS válidos.',1;
            IF EXISTS(SELECT enlace.Id FROM @ListaEnlaces AS enlace WHERE enlace.Id IS NOT NULL GROUP BY enlace.Id HAVING COUNT_BIG(1)>1)
                THROW 51226,N'[HYT_VALIDACION] Un enlace está repetido.',1;
            IF EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NOT NULL AND NOT EXISTS
                (SELECT 1 FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.Id=entrada.Id AND enlace.ClienteId=@Id))
                THROW 51226,N'[HYT_VALIDACION] Un enlace no pertenece al cliente o ya no existe.',1;
        END;
'''

SAVE = '''        IF @Enlaces IS NOT NULL
        BEGIN
            DELETE enlace FROM dbo.ClienteEnlaceDocumento AS enlace WHERE enlace.ClienteId=@Id
                AND NOT EXISTS(SELECT 1 FROM @ListaEnlaces AS entrada WHERE entrada.Id=enlace.Id);
            UPDATE enlace SET Descripcion=entrada.Descripcion,Url=entrada.Url,Orden=entrada.Orden,
                FechaActualizacion=SYSUTCDATETIME(),UsuarioActualizacion=@UsuarioId
            FROM dbo.ClienteEnlaceDocumento AS enlace INNER JOIN @ListaEnlaces AS entrada ON entrada.Id=enlace.Id WHERE enlace.ClienteId=@Id;
            INSERT INTO dbo.ClienteEnlaceDocumento(ClienteId,Descripcion,Url,Orden,UsuarioRegistro)
            SELECT @Id,entrada.Descripcion,entrada.Url,entrada.Orden,@UsuarioId FROM @ListaEnlaces AS entrada WHERE entrada.Id IS NULL;
        END;
'''
