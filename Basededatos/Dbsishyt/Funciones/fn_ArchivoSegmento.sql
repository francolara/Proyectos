-- =============================================
-- Author:        FRANCO LARA
-- Create date:   06/11/2026
-- Description:   Segmento ASCII seguro para carpetas por número de documento, idéntico al generador web.
-- =============================================
USE [SistemaWebHYT];
GO
CREATE OR ALTER FUNCTION dbo.fn_ArchivoSegmento(@Numero NVARCHAR(100))
RETURNS NVARCHAR(100)
AS
BEGIN
    DECLARE @Origen NVARCHAR(100)=LTRIM(RTRIM(@Numero)),@Resultado NVARCHAR(100)=N'',@Indice INT=1,@Caracter NCHAR(1);
    WHILE @Indice<=LEN(@Origen COLLATE Latin1_General_100_BIN2)
    BEGIN
        SET @Caracter=SUBSTRING(@Origen COLLATE Latin1_General_100_BIN2,@Indice,1);
        SET @Resultado=@Resultado+CASE WHEN (@Caracter COLLATE Latin1_General_100_BIN2 LIKE N'[A-Za-z0-9]' OR @Caracter IN(N'_',N'-')) THEN @Caracter ELSE N'_' END;
        SET @Indice=@Indice+1;
    END;
    RETURN @Resultado;
END;
GO
