CREATE FUNCTION dbo.FechaEnLetras (@Fecha DATE)
RETURNS NVARCHAR(MAX)
AS
BEGIN
    DECLARE @Dia INT, @Mes INT, @Ano INT
    DECLARE @DiaSemana NVARCHAR(20), @MesLetras NVARCHAR(20)
    DECLARE @Resultado NVARCHAR(MAX)

    -- Separar el día, mes y año
    SET @Dia = DAY(@Fecha)
    SET @Mes = MONTH(@Fecha)
    SET @Ano = YEAR(@Fecha)

    -- Obtener el día de la semana en letras
    SET @DiaSemana = CASE DATEPART(WEEKDAY, @Fecha)
        WHEN 1 THEN 'domingo'
        WHEN 2 THEN 'lunes'
        WHEN 3 THEN 'martes'
        WHEN 4 THEN 'miércoles'
        WHEN 5 THEN 'jueves'
        WHEN 6 THEN 'viernes'
        WHEN 7 THEN 'sábado'
    END

    -- Convertir mes a letras
    SET @MesLetras = CASE @Mes
        WHEN 1 THEN 'enero'
        WHEN 2 THEN 'febrero'
        WHEN 3 THEN 'marzo'
        WHEN 4 THEN 'abril'
        WHEN 5 THEN 'mayo'
        WHEN 6 THEN 'junio'
        WHEN 7 THEN 'julio'
        WHEN 8 THEN 'agosto'
        WHEN 9 THEN 'septiembre'
        WHEN 10 THEN 'octubre'
        WHEN 11 THEN 'noviembre'
        WHEN 12 THEN 'diciembre'
    END

    -- Formatear la fecha en letras
    SET @Resultado = @DiaSemana + ', ' + CAST(@Dia AS NVARCHAR(2)) + ' de ' + @MesLetras + ' de ' + CAST(@Ano AS NVARCHAR(4))

    RETURN @Resultado
END