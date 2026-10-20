-- =============================================
-- Author:      FRANCO LARA / Codex
-- Create date: 15/10/2026
-- Description: Lista países, departamentos, provincias o distritos para el mantenimiento de ubigeo.
-- =============================================
CREATE OR ALTER PROCEDURE dbo.spu_MantUbigeo_Listar
    @Tipo TINYINT,
    @Busqueda VARCHAR(255) = ''
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        SET @Busqueda = LTRIM(RTRIM(ISNULL(@Busqueda, '')));

        IF @Tipo = 0
        BEGIN
            SELECT P.IdDato AS Codigo, P.IdDato AS IdPais, P.GlsDato AS GlsPais
            FROM dbo.Datos AS P
            WHERE P.IdTipoDatos = '02'
              AND (@Busqueda = '' OR P.GlsDato LIKE '%' + @Busqueda + '%')
            ORDER BY P.GlsDato;
            RETURN;
        END;

        IF @Tipo = 1
        BEGIN
            SELECT D.IdDistrito AS Codigo, P.IdDato AS IdPais, P.GlsDato AS GlsPais,
                   D.IdDpto, D.GlsUbigeo AS GlsDpto
            FROM dbo.Datos AS P
            INNER JOIN dbo.Ubigeo AS D ON D.IdPais = P.IdDato
            WHERE P.IdTipoDatos = '02' AND D.IdProv = '00' AND D.IdDist = '00'
              AND (@Busqueda = '' OR D.GlsUbigeo LIKE '%' + @Busqueda + '%')
            ORDER BY D.GlsUbigeo;
            RETURN;
        END;

        IF @Tipo = 2
        BEGIN
            SELECT R.IdDistrito AS Codigo, P.IdDato AS IdPais, P.GlsDato AS GlsPais,
                   D.IdDpto, D.GlsUbigeo AS GlsDpto, R.IdProv, R.GlsUbigeo AS GlsProv
            FROM dbo.Datos AS P
            INNER JOIN dbo.Ubigeo AS D ON D.IdPais = P.IdDato AND D.IdProv = '00' AND D.IdDist = '00'
            INNER JOIN dbo.Ubigeo AS R ON R.IdPais = D.IdPais AND R.IdDpto = D.IdDpto
                                      AND R.IdProv <> '00' AND R.IdDist = '00'
            WHERE P.IdTipoDatos = '02'
              AND (@Busqueda = '' OR R.GlsUbigeo LIKE '%' + @Busqueda + '%')
            ORDER BY R.GlsUbigeo;
            RETURN;
        END;

        IF @Tipo = 3
        BEGIN
            SELECT I.IdDistrito AS Codigo, P.IdDato AS IdPais, P.GlsDato AS GlsPais,
                   D.IdDpto, D.GlsUbigeo AS GlsDpto, R.IdProv, R.GlsUbigeo AS GlsProv,
                   I.IdDist, I.GlsUbigeo AS GlsDist
            FROM dbo.Datos AS P
            INNER JOIN dbo.Ubigeo AS D ON D.IdPais = P.IdDato AND D.IdProv = '00' AND D.IdDist = '00'
            INNER JOIN dbo.Ubigeo AS R ON R.IdPais = D.IdPais AND R.IdDpto = D.IdDpto
                                      AND R.IdProv <> '00' AND R.IdDist = '00'
            INNER JOIN dbo.Ubigeo AS I ON I.IdPais = R.IdPais AND I.IdDpto = R.IdDpto
                                      AND I.IdProv = R.IdProv AND I.IdDist <> '00'
            WHERE P.IdTipoDatos = '02'
              AND (@Busqueda = '' OR I.GlsUbigeo LIKE '%' + @Busqueda + '%')
            ORDER BY I.GlsUbigeo;
        END;
    END TRY
    BEGIN CATCH
        DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
        SELECT @ErrorMessage=ERROR_MESSAGE(), @ErrorSeverity=ERROR_SEVERITY(), @ErrorState=ERROR_STATE();
        RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
